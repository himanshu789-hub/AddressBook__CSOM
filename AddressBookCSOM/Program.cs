using System;
using System.Collections.Generic;
using AddressBook.Contracts;
using AddressBook.Models;
using AddressBook.Interfaces;
using System.Runtime.Remoting;
using System.Linq;
using ConsoleTables;
using System.Text;
using AddressBook.Exception;
using System.IO;

namespace AddressBook.Console
{
    class Program
    {
        
        static void Main(string[] args)
        {
            
            AddressBookCSOM Controller = null ;
            List<User> users = new List<User>();
            try
            {
                System.Console.WriteLine("Processing.Please Wait . . .");
                Controller = new AddressBookCSOM();
                if (!Controller.DoCategoryFieldAdded)
                    Controller.AddTaxonomyField("Department", "Department", "Department", "Field", new string[] { "IT", "Support", "Sales" });
                if (!Controller.DoCoordinatorFieldAdded)
                    Controller.AddPersonOrGroupField("Coordinator", "Coordinator", true);
                users.AddRange(Controller.GetAllSiteUser());
            }
            catch(ServerException)
            {
                System.Console.WriteLine("Server Error . . .");
            }
            bool flag = true; 
            System.Console.WriteLine("Hello User . . .");
            while (flag)
            {
                System.Console.WriteLine("## Please select any option . . .");
                System.Console.WriteLine("## 1.Add\t2.Delete\t3.Update\t4.View All");
                System.Console.WriteLine("Your Choice := ");
                char key =(char) System.Console.Read();
                System.Console.ReadLine();
                switch (key)
                {
                    case '1':
                        {
                            AddressDetail Item = new AddressDetail();
                            System.Console.WriteLine("Enter Full Name : ");
                            Item.FullName = System.Console.ReadLine();
                            System.Console.WriteLine("Enter Contact Number : ");
                            Item.PhoneNumber = System.Console.ReadLine();
                            System.Console.WriteLine("Enter Address : ");
                            Item.Address = System.Console.ReadLine();
                            
                            System.Console.WriteLine("Choose Coordinator : ");
                            System.Console.WriteLine("Processing Wait . . .");
                            DisplayAllPeople(users);
                            System.Console.WriteLine("Enter Space Seperated Ids : ");
                            int[] ids = null;
                            while (true)
                            {
                                try
                                {

                                    ids = Array.ConvertAll( System.Console.ReadLine().Trim().Split(new char[] { ' ' }),s=>int.Parse(s));
                                    bool ValidIds = true;
                                    foreach (int id in ids)
                                    {
                                        if (!users.Select(e => e.Id).Contains(id))
                                            ValidIds = false;
                                    }
                                    if (ValidIds)
                                        break;
                                    System.Console.WriteLine("Please, Enter Valid Ids . . .");
                                }
                                catch(System.Exception)
                                {
                                    System.Console.WriteLine("Please, Enter Only Space Seperated Number . . .");
                                }
                            }
                            Item.Users = users.FindAll(e => ids.Contains(e.Id));
                            System.Console.WriteLine("Enter Department of Person : ");
                            System.Console.WriteLine("Processing.Please Wait . . .");
                            List<Term> terms = Controller.GetTermAssociatedWithTaxonomyField();
                            int count = terms.Count;
                            DisplayDepartments(terms);
                            System.Console.WriteLine("Your Choice := ");
                            int choice;
                            while (true)
                            {
                                choice = int.Parse(System.Console.ReadLine());
                                if (choice > 0 && choice <= count)
                                    break;
                                System.Console.WriteLine("Please,  Enter Valid Choice . . .\nYour Choice := ");
                            }
                            Term term = terms.ElementAt(choice - 1);
                            Item.Term = term;
                            
                            try
                            {
                                System.Console.WriteLine("Processing.Please Wait . . .");
                                Controller.Create(Item);
                                System.Console.WriteLine("Added Successfully.");
                            }
                            catch (ServerException)
                            {
                                System.Console.WriteLine("Sorry, Connection Error . . .\nPlese Try Again");
                            }
                        }
                        break;
                    case '2':
                        {
                            try
                            {
                                System.Console.WriteLine("Processing.Please Wait . . .");
                                List<AddressDetail> Addresses = Controller.GetAll();
                                DisplayRecords(Addresses);
                                System.Console.WriteLine("Enter User Id To Delete := ");
                                int choice = int.Parse(System.Console.ReadLine());
                                ValidateId(choice, Addresses);
                                System.Console.WriteLine("Processing.Please Wait . . .");
                                Controller.Delete(choice);
                            }
                            catch(InvalidOperationException)
                            {
                                System.Console.WriteLine("Please,Enter Valid Id.Try Again . . .");
                            }
                            catch(System.Exception e)
                            {
                                System.Console.WriteLine("Sorry, Connection Error . . .\nPlese Try Again . . .");
                            }
                        }
                        break;
                    case '3':
                        {
                            try
                            {
                                System.Console.WriteLine("Processing.Please Wait . . .");
                                List<AddressDetail> Addresses = Controller.GetAll();
                                DisplayRecords(Addresses);
                                System.Console.WriteLine("Enter Person User Id := ");
                                int SelectedUserId = int.Parse(System.Console.ReadLine());
                                AddressDetail OldAddressDetail = ValidateId(SelectedUserId, Addresses);
                                System.Console.WriteLine("Change --> \n1.Full Name\t2.Phone Number\t3.Address\t4.Co-ordinators\t5.Department\t6.All User Detail");
                                System.Console.WriteLine("Your Choice := ");
                                int ChooseUpdateOption = System.Console.Read()-'0';
                                System.Console.ReadLine();
                                AddressDetail NewAddressDetail = OldAddressDetail;
                                switch(ChooseUpdateOption)
                                {
                                    case 1:
                                        System.Console.WriteLine("Enter New Full Name : ");
                                     
                                        NewAddressDetail.FullName =  System.Console.ReadLine();
                                        break;
                                    case 2:
                                        System.Console.WriteLine("Enter New Phone Number : ");
                                        NewAddressDetail.PhoneNumber = System.Console.ReadLine();
                                        break;
                                    case 3:
                                        System.Console.WriteLine("Enter New Address : ");
                                        NewAddressDetail.Address = System.Console.ReadLine();
                                        break;
                                    case 4:
                                        {
                                            DisplayAllPeople(users);
                                            System.Console.WriteLine("Enter Space Seperated Ids : ");
                                            int[] ids = null;
                                            while (true)
                                            {
                                                try
                                                {

                                                    ids =Array.ConvertAll(System.Console.ReadLine().Trim().Split(new char[] { ' ' }),s=>int.Parse(s));
                                                    bool ValidIds = true;
                                                    foreach (int id in ids)
                                                    {
                                                        if (!users.Select(e => e.Id).Contains(id))
                                                        {
                                                            ValidIds = false;
                                                            break;
                                                        }
                                                    }
                                                    if (ValidIds)
                                                        break;
                                                    System.Console.WriteLine("Please, Enter Valid Ids . . .");
                                                }
                                                catch (System.Exception)
                                                {
                                                    System.Console.WriteLine("Please, Enter Only Space Seperated Number . . .");
                                                }
                                            }
                                            NewAddressDetail.Users.Clear();
                                            for (int i=0;i<ids.Length;i++)
                                            {
                                              
                                                User user = new User() { Id = ids[i] };
                                                NewAddressDetail.Users.Add(user);
                                            }
                                        }
                                        break;
                                    case 5:
                                        {
                                            System.Console.WriteLine("Processing.Please Wait . . .");
                                            List<Term> terms = Controller.GetTermAssociatedWithTaxonomyField();
                                            DisplayDepartments(terms);
                                            int count = terms.Count;
                                            
                                            System.Console.WriteLine("Your Choice := ");
                                            int choice;
                                            while (true)
                                            {
                                                choice = int.Parse(System.Console.ReadLine());
                                                if (choice > 0 && choice <= count)
                                                    break;
                                                System.Console.WriteLine("Please,  Enter Valid Choice . . .\nYour Choice := ");
                                            }
                                            Term term = terms.ElementAt(choice - 1);
                                            NewAddressDetail.Term = term;

                                        }
                                        break;
                                    case 6:
                                        {
                                          
                                            System.Console.WriteLine("Enter Full Name : ");
                                            NewAddressDetail.FullName = System.Console.ReadLine();
                                            System.Console.WriteLine("Enter Contact Number : ");
                                            NewAddressDetail.PhoneNumber = System.Console.ReadLine();
                                            System.Console.WriteLine("Enter Address : ");
                                            NewAddressDetail.Address = System.Console.ReadLine();

                                            System.Console.WriteLine("Choose Coordinator : ");
                                            System.Console.WriteLine("Processing Wait . . .");
                                            DisplayAllPeople(users);
                                            System.Console.WriteLine("Enter Space Seperated Ids : ");
                                            int[] ids = null;
                                            while (true)
                                            {
                                                try
                                                {

                                                    ids = Array.ConvertAll(System.Console.ReadLine().Trim().Split(new char[] { ' ' }),s=>int.Parse(s));
                                                    bool ValidIds = true;
                                                    foreach (int id in ids)
                                                    {
                                                        if (!users.Select(e => e.Id).Contains(id))
                                                            ValidIds = false;
                                                    }
                                                    if (ValidIds)
                                                        break;
                                                    System.Console.WriteLine("Please, Enter Valid Ids . . .");
                                                }
                                                catch (System.Exception)
                                                {
                                                    System.Console.WriteLine("Please, Enter Only Space Seperated Number . . .");
                                                }
                                            }
                                            NewAddressDetail.Users.Clear();
                                            for (int i=0;i<ids.Length;i++)
                                            {
                                                
                                                User user = new User() { Id = ids[i] };
                                                NewAddressDetail.Users.Add(user);
                                            }
                                            System.Console.WriteLine("Enter Department of Person : ");
                                            System.Console.WriteLine("Processing.Please Wait . . .");
                                            List<Term> terms = Controller.GetTermAssociatedWithTaxonomyField();
                                            int count = terms.Count;
                                            DisplayDepartments(terms);
                                            System.Console.WriteLine("Your Choice := ");
                                            int choice;
                                            while (true)
                                            {
                                                choice = int.Parse(System.Console.ReadLine());
                                                if (choice > 0 && choice <= count)
                                                    break;
                                                System.Console.WriteLine("Please,  Enter Valid Choice . . .\nYour Choice := ");
                                            }
                                            Term term = terms.ElementAt(choice - 1);
                                            NewAddressDetail.Term = term;

                                        }
                                        break;

                                    default:
                                        System.Console.WriteLine("Invalid User Id.Please Try Again . . .");
                                        break;
                                }
                                System.Console.WriteLine("Processing.Please Wait . . .");
                                Controller.Update(SelectedUserId,NewAddressDetail);
                                System.Console.WriteLine("< < < Updated SuccessFully > > >");
                            }
                            catch (InvalidUserIdException)
                            {
                                System.Console.WriteLine("Please, Enter Valid User Id");
                            }
                            catch(System.Exception)
                            {
                                System.Console.WriteLine("Sorry, Server Error.Please Try Again...");
                            }
                        }
                        break;
                    case '4':
                        {

                            try
                            {
                                System.Console.WriteLine("Processing.Please Wait . . .");
                                List<AddressDetail> Addresses = Controller.GetAll();
                                DisplayRecords(Addresses);
                            }
                            catch (ServerException e)
                            {
                                System.Console.WriteLine("Sorry Server Error . . .\nTrying Again ");
                            }

                        }
                        break;
                    case 'e':
                        flag = false;
                        break;
                    default:
                        System.Console.WriteLine("Please, Enter Valid Option . . .");
                        break;
                }
            }

        }
        static void DisplayRecords(List<AddressDetail> Addresses)
        {
            System.Console.WriteLine("All Person Details : ");
            var table = new ConsoleTable("Id", "Full Name", "Phone Number", "Address", "Department", "Co-ordinator");
            for (int i = 0; i < Addresses.Count; i++)
            {
                StringBuilder coordinatorBuilder = new StringBuilder();
                for(int j=0;j<Addresses[i].Users.Count;j++)
                {
                    coordinatorBuilder.Append(" " + (j + 1) + "." + Addresses[i].Users[j].FullName);
                }
                table.AddRow(Addresses[i].Id, Addresses[i].FullName, Addresses[i].PhoneNumber,Addresses[i].Address ,Addresses[i].Term.Title, coordinatorBuilder.ToString());
            }
            table.Write();
        }
        static AddressDetail ValidateId(int Id,List<AddressDetail> Addresses)
        {
            try
            {
                return Addresses.First(e => e.Id == Id);
            }
            catch(InvalidOperationException)
            {
                throw new InvalidUserIdException();
            }
        }
       static void DisplayAllPeople(List<User> users)
        {
            System.Console.WriteLine("All User : ");
            var table = new ConsoleTable("Id","Full Name");
            foreach (User user in users)
                table.AddRow(user.Id,user.FullName);
            table.Write();
       }
        static void DisplayDepartments(List<Term> terms)
        {
            System.Console.WriteLine("Departments are : ");
            var table = new ConsoleTable("S.No.", "Title");
            for (int i = 0; i < terms.Count; i++)
                table.AddRow(i + 1, terms[i].Title);
            table.Write();
        }
    }
}
