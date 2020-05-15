using System;
using System.Linq;
using System.Configuration;
using Collection = System.Collections.Generic;
using Microsoft.SharePoint.Client.Taxonomy;
using SiteUser = AddressBook.Models.User;
using User = Microsoft.SharePoint.Client.User;
using TermValue = AddressBook.Models.Term;
using Term = Microsoft.SharePoint.Client.Taxonomy.Term;
using Microsoft.SharePoint.Client;
using AddressBook.Interfaces;
using AddressBook.Models;
using AddressBook.Constants;
namespace AddressBook.Contracts
{
    public class AddressBookCSOM : IAddressBook
    {
       bool IsCoordinatorAdded;
       bool IsCategoryFieldAdded;
       public bool DoCoordinatorFieldAdded { get { return IsCoordinatorAdded; } }
       public bool DoCategoryFieldAdded{  get { return IsCategoryFieldAdded; }}

      readonly  string FullNameInternalName = "Title";
      readonly  string PrimaryIndex = "ID";
      readonly  string PhoneNumberInternalName = "Phone_x0020_No_x002e_";
      readonly  string AddressInternalName = "Address";
      readonly  string CoordinatorInternalName = "Coordinator";
      readonly string DepartmentInternalName = "Department";

     readonly string ViewName = ConfigurationManager.AppSettings[Constants.Constants.AppSettingConstants.ViewName];
     readonly string TargetList = ConfigurationManager.AppSettings[Constants.Constants.AppSettingConstants.ListTitle];

        ClientContext context;
        public AddressBookCSOM(){
            context = new ClientContext(ConfigurationManager.ConnectionStrings[Constants.Constants.ConnectionConstants.SiteURL].ConnectionString);
            string Password = ConfigurationManager.AppSettings[Constants.Constants.AppSettingConstants.Password];
            System.Security.SecureString SecureString = new System.Security.SecureString();
            Password.ToList().ForEach(SecureString.AppendChar);
            context.Credentials = new SharePointOnlineCredentials(ConfigurationManager.AppSettings[Constants.Constants.AppSettingConstants.UserId],SecureString);
            List list = context.Web.Lists.GetByTitle(TargetList);
            context.Load(list.Fields);
            context.ExecuteQuery();
            Collection.List<Field> ListColumns =  list.Fields.ToList();
            IsCategoryFieldAdded = ListColumns.FirstOrDefault(e => e.Title == ConfigurationManager.AppSettings[Constants.Constants.AppSettingConstants.DepartmentColumn]) != null;
            IsCoordinatorAdded = ListColumns.FirstOrDefault(e => e.Title == ConfigurationManager.AppSettings[Constants.Constants.AppSettingConstants.CoordinatorColumn]) != null;
        }
        public AddressDetail Create(AddressDetail Item)
        {
            List AddressList = context.Web.Lists.GetByTitle(TargetList);
            FieldUserValue[] fieldUserValueCollection = new FieldUserValue[Item.Users.Count];
            for(int i=0;i<Item.Users.Count;i++)
              fieldUserValueCollection[i] = new FieldUserValue() { LookupId = Item.Users[i].Id };
            context.Load(AddressList);
            context.ExecuteQuery();
            Field field = AddressList.Fields.GetByTitle(DepartmentInternalName);
            context.Load(field);
            context.ExecuteQuery();
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem newItem = AddressList.AddItem(itemCreateInfo);

            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            TaxonomyFieldValue taxonomyFieldValue = new TaxonomyFieldValue() { Label = Item.Term.Title, TermGuid = Item.Term.Id.ToString() };
            taxonomyFieldValue.WssId = -1;
            taxonomyField.SetFieldValueByValue(newItem, taxonomyFieldValue);
            taxonomyField.Update();
            //newItem[DepartmentInternalName] = 
          
            newItem[CoordinatorInternalName] = fieldUserValueCollection;
            newItem[FullNameInternalName] = Item.FullName;
            newItem[AddressInternalName] = Item.Address;
            newItem[PhoneNumberInternalName] = Item.PhoneNumber;
            newItem.Update();
            AddressList.Update();
            context.ExecuteQuery();
            Item.Id = newItem.Id;
            context.Load(newItem);
            context.ExecuteQuery();
            return Item;
        }
        public AddressDetail Update(int Id, AddressDetail Item)
        {
            List AddressList = context.Web.Lists.GetByTitle(TargetList);
            context.Load(AddressList);
            context.ExecuteQuery();
            Field field = AddressList.Fields.GetByTitle(DepartmentInternalName);
            context.Load(field);
            context.ExecuteQuery();

            ListItem listItem = AddressList.GetItemById(Id);
            context.Load(listItem);
            context.ExecuteQuery();

            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            TaxonomyFieldValue taxonomyFieldValue = new TaxonomyFieldValue() { Label = Item.Term.Title, TermGuid = Item.Term.Id.ToString() };
            taxonomyFieldValue.WssId = -1;
            taxonomyField.SetFieldValueByValue(listItem, taxonomyFieldValue);
            taxonomyField.Update();

            listItem[FullNameInternalName] = Item.FullName;
            listItem[AddressInternalName] = Item.Address;
            listItem[PhoneNumberInternalName] = Item.PhoneNumber;
            FieldUserValue[] fieldUserValues = new FieldUserValue[Item.Users.Count];
            for (int i = 0; i < Item.Users.Count; i++)
                fieldUserValues[i] = new FieldUserValue() { LookupId = Item.Users[i].Id };
            listItem[CoordinatorInternalName] = fieldUserValues;

            //listItem[CoordinatorInternalName] = fieldUserValues;
            listItem.Update();
            context.ExecuteQuery();
            return Item;
        }
        public void Delete(int Id)
        {
            List AddressList = context.Web.Lists.GetByTitle(TargetList);
            ListItem listItem = AddressList.GetItemById(Id);
            listItem.DeleteObject();
            AddressList.Update();
            context.ExecuteQuery();
        }

        public Collection.List<AddressDetail> GetAll()
        {
            Collection.List<AddressDetail> Addresses = new Collection.List<AddressDetail>();
            List AddressList = context.Web.Lists.GetByTitle(TargetList);
            context.Load(AddressList);
            context.ExecuteQuery();
            View View = AddressList.Views.GetByTitle(ViewName);
            context.Load(View);
            context.ExecuteQuery();
            CamlQuery Query = new CamlQuery();
            Query.ViewXml = View.ViewQuery;
            ListItemCollection ListItems = AddressList.GetItems(Query);
            context.Load(ListItems);
            context.ExecuteQuery();
            foreach (ListItem Item in ListItems)
            {
                context.Load(Item);
                context.ExecuteQuery();

                TaxonomyFieldValue taxonomyField = Item[DepartmentInternalName] as TaxonomyFieldValue;
                FieldUserValue[] fieldUserValues = Item[CoordinatorInternalName] as FieldUserValue[];
                SiteUser[] users = null;
                if (fieldUserValues!=null)
                {
                    users = new SiteUser[fieldUserValues.Length];
                    for (int i = 0; i < fieldUserValues.Length; i++)
                    {
                        FieldUserValue fieldUserValue = fieldUserValues[i];
                        User user = context.Web.SiteUsers.GetById(fieldUserValue.LookupId);
                        context.Load(user);
                        context.ExecuteQuery();
                        SiteUser siteUser = new SiteUser() { FullName = user.Title, Id = user.Id };
                        users[i] = siteUser;
                    }
                }

                Addresses.Add(new AddressDetail() {Users=users.ToList(),Term=new TermValue() {Id=Guid.Parse(taxonomyField.TermGuid),Title=taxonomyField.Label }, Address = Item[AddressInternalName].ToString(), FullName = Item[FullNameInternalName].ToString(), Id = int.Parse(Item[PrimaryIndex].ToString()), PhoneNumber = Item[PhoneNumberInternalName].ToString() });
            }
            return Addresses;
        }

      
        
        public bool AddPersonOrGroupField(string displayName,string internalName,bool isMultiValue)
        {
            List list = context.Web.Lists.GetByTitle(TargetList);
            context.Load(list);
            context.ExecuteQuery();
            FieldCollection fields = list.Fields;
            context.Load(fields);
            context.ExecuteQuery();
            string[] fieldsInternalName = fields.Select(e => e.InternalName).ToArray();
            if(fieldsInternalName.Contains(internalName))
                throw new Microsoft.SharePoint.Client.ClientRequestException("Field Already Exists");
             
            string userFieldSchema = "<Field Type='User' DisplayName='"+displayName+"' Name='"+internalName+"' StaticName='"+
                internalName+"' UserSelectionMode='PeopleOnly' Mult='"+(isMultiValue?"TRUE":"FALSE")+"'/>";
            fields.AddFieldAsXml(userFieldSchema, true, AddFieldOptions.AddFieldInternalNameHint);  
            context.ExecuteQuery();
            return true;
        }
        public bool AddTaxonomyField(string displayName,string internalName,string groupName,string termsetName,string[] terms)
        {
            if(internalName.ToCharArray().Contains(' '))
                throw new Microsoft.SharePoint.Client.ClientRequestException("Field Internal Name Not Valid");

            List list = context.Web.Lists.GetByTitle(TargetList);
            context.Load(list.Fields);
            context.ExecuteQuery();
            string[] fieldsInternalName = list.Fields.Select(e => e.InternalName).ToArray();
            if (fieldsInternalName.Contains(internalName))
                throw new Microsoft.SharePoint.Client.ClientRequestException("Field Already Exists");

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            taxonomySession.UpdateCache();
            context.Load(taxonomySession.TermStores);
            context.ExecuteQuery();
            TermStore termStore = taxonomySession.TermStores.ToList().FirstOrDefault(e=>e.Name==ConfigurationManager.AppSettings[Constants.Constants.AppSettingConstants.MedataServiceName]);
            if (termStore == null)
                throw new Microsoft.SharePoint.Client.ClientRequestException("Invalid Metadata Service Name");
            context.Load(termStore);
            context.ExecuteQuery();
            
            TermGroupCollection groups = termStore.Groups;
            context.Load(groups);
            TermGroup group = groups.GetByName(groupName);
            context.Load(group);
            context.ExecuteQuery();
            Guid termSetId = Guid.Empty;
            if (group==null)
            {
                TermGroup termGroup = termStore.CreateGroup(groupName,Guid.NewGuid());
                context.Load(termGroup);
                context.ExecuteQuery();
                TermSet termSet = termGroup.CreateTermSet(termsetName, Guid.NewGuid(), termStore.DefaultLanguage);
                foreach (string term in terms)
                {
                    termSet.CreateTerm(term, termStore.DefaultLanguage, Guid.NewGuid());
                }
                context.Load(termSet);
                context.ExecuteQuery();
                termSetId = termSet.Id;
                }
            else
            {
                context.Load(group.TermSets);
                context.ExecuteQuery();
                TermSet termSet = group.TermSets.ToList().FirstOrDefault(e=>e.Name==termsetName);
                if (termSet!=null)
                {
                    TermCollection terms1 = termSet.GetAllTerms();
                    foreach (string term in terms)
                    {
                        termSet.CreateTerm(term, termStore.DefaultLanguage, Guid.NewGuid());
                    }
                    context.Load(termSet);
                    context.ExecuteQuery();
                    termSetId = termSet.Id;
                }
                else 
                {
                   TermSet termSet1 = group.CreateTermSet(termsetName, Guid.NewGuid(),termStore.DefaultLanguage);
                    foreach (string term in terms)
                    {
                        termSet1.CreateTerm(term, termStore.DefaultLanguage, Guid.NewGuid());
                    }
                    context.Load(termSet1);
                    context.ExecuteQuery();
                    termSetId = termSet1.Id;
                }
            }

            string schemaTaxonomyField = "<Field  Type='TaxonomyFieldType' Name='" + internalName + "' " +
             "StaticName='" + internalName + "' DisplayName = '" + displayName + "' /> ";

            Field field = list.Fields.AddFieldAsXml(schemaTaxonomyField, true, AddFieldOptions.AddFieldInternalNameHint);
            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStore.Id;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.Update();
            context.ExecuteQuery();
            return true;
        }
        public Collection.List<SiteUser> GetAllSiteUser()
        {
            context.Load(context.Web.SiteUsers);
            context.ExecuteQuery();
            UserCollection users = context.Web.SiteUsers;
            Collection.List < SiteUser > siteUsers = new Collection.List<SiteUser>(); 
            foreach(User user in users)
            {
                siteUsers.Add(new SiteUser() { FullName = user.UserPrincipalName, Id = user.Id });
            }
            return siteUsers;
        }
        public Collection.List<TermValue> GetTermAssociatedWithTaxonomyField()
        {
            Collection.List<TermValue> termValues = new Collection.List<TermValue>();
            List list =   context.Web.Lists.GetByTitle(TargetList);
            context.Load(list.Fields);
            context.ExecuteQuery();
            Field field = list.Fields.GetByTitle(DepartmentInternalName);
            context.Load(field);
            context.ExecuteQuery();
            TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(field);
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            context.Load(taxonomySession);
            context.ExecuteQuery();
            TermStore store = taxonomySession.TermStores.GetByName(ConfigurationManager.AppSettings["Metadata Service Name"]);
            context.Load(store);
            context.ExecuteQuery();
            TermSet termSet = store.GetTermSet(taxonomyField.TermSetId);
            context.Load(termSet.Terms);
            context.ExecuteQuery();
            foreach(Term term in termSet.Terms)
            {
                termValues.Add(new TermValue() { Id = term.Id, Title = term.Name });
            }
            return termValues;
        }
    }
}
