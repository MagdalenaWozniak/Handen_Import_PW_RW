using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ONE032
{
    public struct ONE032Data
    {
        public string PracKod { set; get; }     //Akronim                  DOH DOH_DodId -> DOD DOD_DodId  DOD_PraId-> PRA_PraId 
        public string DOH_Nazwa { set; get; }    //Nazwa Składnika
        public DateTime DOH_WaznyOd { set; get; }

        public DateTime DOH_WaznyDo { set; get; }

        public decimal DOH_Wartosc1 { set; get; }
        //public string ListaPlac { set; get; }    //symbol tyup skladnika
        public string NazwaBazy         {set;get;}
        public string NumerZlecenia     {set;get;}
        public string Material          {set;get;}
        public string Odbiorca          {set;get;}
        public string MiejsceDostawy    {set;get;}
        public string NrDowoduWydania   {set;get;}
        public double IloscWM15         {set;get;}
        public decimal Cena              {set;get;}
        public decimal WartoscNetto      {set;get;}
        public decimal KwotaPodatku      {set;get;}
        public decimal WartoscBrutto     { set; get; }
    }
}
