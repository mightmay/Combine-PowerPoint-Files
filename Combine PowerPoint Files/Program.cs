using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenXmlPowerTools;
using System.Web;
using System.IO;
using ns;

namespace CombinePPTXfiles
{
    class Program
    {
        static void Main(string[] args)
        {
            string currentdirectory = Directory.GetCurrentDirectory();
            string[] listofppt = Directory.GetFiles(currentdirectory, "*.pptx");
            NumericComparer ns = new NumericComparer();
            Array.Sort(listofppt, ns);

            Console.WriteLine("The current directory is {0}", currentdirectory);
            string savename = "COMBINED.pptx";
            List<SlideSource> sources = new List<SlideSource>();
            if (listofppt.Length > 0)
            {
                for (int i = 0; i < listofppt.Count(); i++)
                {
                    sources.Add(new SlideSource(new PmlDocument(listofppt[i]), true));

                }
                PresentationBuilder.BuildPresentation(sources, savename);
            }

        }
    }
}


