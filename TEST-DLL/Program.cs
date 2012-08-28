using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace TEST_DLL
{
    class Program
    {
        static void Main(string[] args)
        {
            SDNSharePoint.EasySPList myESP = new SDNSharePoint.EasySPList("http://painsp");
            myESP.CleanRetrunValues = true;
            DataSet ds = null;

            //DataSet ds = myESP.SearchList("TestMe");
            //DataSet ds =  myESP.SearchList("TestMe", null, "Title;User", ';');

            //string RetrunINof = myESP.AddARowToList("TestMe", "title=program,user=program,emailaddress=shane@program.com,JustANumber=2112,YesOrNo=0,Space Me=I like to do things like this.", ',');
            string RetrunINof = myESP.AddARowToList("TestMe", "Title=program,User=program,Space Me=I like to do things like this.,emailaddress=shane@program.com,JustANumber=2112,YesOrNo=0", ',');

            Console.WriteLine(RetrunINof);


            foreach (DataRow dr in ds.Tables["TestMe"].Rows)
            {
                foreach (DataColumn dc in ds.Tables["TestMe"].Columns)
                {
                    Console.WriteLine("Col - {0}   Data - {1}",dc.ColumnName,dr[dc.ColumnName]);
                }


            }

        }
    }
}
