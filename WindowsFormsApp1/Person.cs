using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class Person
    {

        public Person(string name, int tickets)
        {
            this.Name = name;
            this.Tickets = tickets;
        }

        public string Name
        {
            get;
            set;
        }

        public int Tickets
        {
            get;
            set;
        }

        public Person()
        {
            Name = "";
            Tickets = 0;
        }
    }
}
