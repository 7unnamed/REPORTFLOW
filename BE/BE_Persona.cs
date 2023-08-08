using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BE
{
    public class BE_Persona
    {
        public BE_Persona()
        {

            ListHorario = new List<BE_Horario>();
        }
        public string dni { get; set; }
        public string nombre { get; set; }
        public List<BE_Horario> ListHorario { get; set; }
    }
}
