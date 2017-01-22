using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using tomenglertde.Wax.Model.Wix;

namespace WixRememberPropertiesEditor
{
    public class WixPropertyModel
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public WixPropertyModel()
        {

        }

        public WixPropertyModel(string name, string value)
        {
            Name = name;
            Value = value;
        }

        public WixProperty GetProperty()
        {
            return new WixProperty(Name.ToUpper(), Value);
        } 
    }
}
