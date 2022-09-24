using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CIS_Module_Import.Controller
{
    internal class DbConnection
    {
        public static Model.CISEntities3 DbContext;
        public static Model.CISEntities3 GetContext()
        {
            if (DbContext == null)
            {
                DbContext = new Model.CISEntities3();
            }
            return DbContext;
        }
    }
}
