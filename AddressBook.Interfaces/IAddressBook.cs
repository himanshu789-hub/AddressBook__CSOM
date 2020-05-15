using System.Collections.Generic;
using AddressBook.Models;

namespace AddressBook.Interfaces
{
    public interface IAddressBook
    {
        AddressDetail Create(AddressDetail Item);
        AddressDetail Update(int Id, AddressDetail Item);
        void Delete(int Id);
        List<AddressDetail> GetAll();
        List<Term> GetTermAssociatedWithTaxonomyField();
        List<User> GetAllSiteUser();
    }
}
