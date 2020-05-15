using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddressBook.Exception
{
    public class InvalidUserIdException:System.Exception
    {
       public InvalidUserIdException() :base("Invalid User Id")
        {

        }
    }
}
