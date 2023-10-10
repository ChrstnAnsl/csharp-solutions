using System;
using System.Xml;

class Program
{
    static void Main(string[] args)
    {
        Program.GetDataFromXml("homepage", "header");
    }

    static void GetDataFromXml(String page, String element)
    {
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load("locators.xml");
        XmlNode pageElement = xmlDoc.SelectSingleNode("//" + page);

        foreach (XmlNode node in pageElement.ChildNodes)
        {
            if (node.Name == element)
            {
                Console.WriteLine("Page is: " + node.Name + " and the Element is: " + node.InnerText);
            }
        }
    }
}
