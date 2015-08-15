using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QRCodeGenerator
{
    public class Account
    {
        public string Bank { get; set; }


        public String ReqId { get; set; }
        public string AccountNo { get; set; }
        public string Name { get; set; }
        public int Qty { get; set; }
        public string Branch { get; set; }

        public string Currency { get; set; }
        
        public string StartStNo { get; set; }
        public string EndStNo { get; set; }
        public string Routing { get; set; }
        public string MICRAc { get; set; }
        public string Type { get; set; }

        public Bitmap Image { get; set; }

    }
}
