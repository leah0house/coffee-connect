
using Microsoft.CSharp.RuntimeBinder;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MonthlyEmailer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool debugFlag = false;

            string inputFilePath = "path\\file.csv";
            string archiveFolder = "folderpath";

            Cursor.Current = Cursors.WaitCursor;
            Regex regex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
            List<Person> people = new List<Person>();
            try
            {
                foreach (string readAllLine in File.ReadAllLines(inputFilePath))
                {
                    string[] strArray = regex.Split(readAllLine);
                    people.Add(new Person()
                    {
                        Name = strArray[0],
                        Company = strArray[1],
                        Title = strArray[2],
                        Email = strArray[3]
                    });
                }
            }
            catch
            {
                int num = (int)MessageBox.Show("Error reading file. Check formatting");
                return;
            }
            if ((uint)(people.Count % 2) > 0U)
            {
                int num1 = (int)MessageBox.Show("Odd number of people. Please add or remove user before continuing.");
            }
            else
            {
                List<ArchivePeoplePairs> archivePeoplePairsList = new List<ArchivePeoplePairs>();
                try
                {
                    foreach (string file in Directory.GetFiles(archiveFolder))
                    {
                        ArchivePeoplePairs archivePeoplePairs = new ArchivePeoplePairs();
                        archivePeoplePairs.PeoplePairs = new List<PeoplePair>();
                        foreach (string readAllLine in File.ReadAllLines(file))
                        {
                            PeoplePair peoplePair = new PeoplePair();
                            string[] strArray = regex.Split(readAllLine);
                            peoplePair.Person1 = new Person()
                            {
                                Name = strArray[0]
                            };
                            peoplePair.Person2 = new Person()
                            {
                                Name = strArray[1]
                            };
                            archivePeoplePairs.PeoplePairs.Add(peoplePair);
                        }
                        archivePeoplePairsList.Add(archivePeoplePairs);
                    }
                }
                catch
                {
                    int num2 = (int)MessageBox.Show("Error reading archives. Check formatting for files in " + archiveFolder + ".");
                    return;
                }
                List<PeoplePair> allArchives = new List<PeoplePair>();
                foreach (ArchivePeoplePairs archivePeoplePairs in archivePeoplePairsList)
                    allArchives.AddRange((IEnumerable<PeoplePair>)archivePeoplePairs.PeoplePairs);
                bool failed = false;
                Person stragler = (Person)null;
                List<PeoplePair> peoplePairList = this.RandomizePeople(people, allArchives, ref failed, ref stragler, archiveFolder);
                if (!failed)
                {
                    var olApplication = new Microsoft.Office.Interop.Outlook.Application();
                    string str1 = DateTime.Now.ToString("MMMM");
                    string str2 = DateTime.Now.Year.ToString();
                    List<string> stringList = new List<string>();
                    foreach (PeoplePair peoplePair in peoplePairList)
                    {
                        stringList.Add(peoplePair.Person1.Name + "," + peoplePair.Person2.Name);

                        Outlook.MailItem mailItem = (Outlook.MailItem)
                            olApplication.CreateItem(Outlook.OlItemType.olMailItem);

                        mailItem.To = peoplePair.Person1.Email + "; " + peoplePair.Person2.Email;
                        mailItem.Subject = "Subject- " + str1 + " " + str2;
                        if (debugFlag)
                            mailItem.BCC = "email";
                        else
                            mailItem.BCC = "email";

                        Assembly executingAssembly = Assembly.GetExecutingAssembly();
                        string name = "MonthlyEmailer.Resources.CoffeeConnect.template.html";
                        string str3 = "";
                        using (Stream manifestResourceStream = executingAssembly.GetManifestResourceStream(name))
                        {
                            using (StreamReader streamReader = new StreamReader(manifestResourceStream))
                                str3 = streamReader.ReadToEnd();
                        }
                        string str4 = str3.Replace("Person1", peoplePair.Person1.Name).Replace("Person2", peoplePair.Person2.Name);
                        mailItem.HTMLBody = str4;
                        try
                        {
                            //if (!debugFlag)
                            //{
                                mailItem.Send();
                            //}
                        }
                        catch
                        {
                            int num2 = (int)MessageBox.Show("Error sending to " + peoplePair.Person1.Email + " and " + peoplePair.Person2.Email);
                        }
                    }
                    if (stragler != null)
                    {
                        stringList.Add(stragler.Name + ",");
                        int num2 = (int)MessageBox.Show(stragler.Name + " did not get a match this month.");
                    }
                    File.WriteAllLines(archiveFolder + "\\" + str1 + str2 + ".csv", (IEnumerable<string>)stringList);
                    Process.Start(archiveFolder + "\\"+ str1 + str2 + ".csv");
                    int num3 = (int)MessageBox.Show("Complete!");
                }
                Cursor.Current = Cursors.Default;
            }
        }


        private List<PeoplePair> CreatePeoplePairList(
          List<Person> people,
          ref Person stragler)
        {
            List<PeoplePair> peoplePairList = new List<PeoplePair>();
            Random random = new Random();
            List<int> source = new List<int>();
            while (source.Count < people.Count)
            {
                int num = random.Next(people.Count);
                if (!source.Contains(num))
                    source.Add(num);
            }
            List<int> list1 = source.Take<int>(source.Count / 2).ToList<int>();
            List<int> list2 = source.Skip<int>(source.Count / 2).ToList<int>();
            int count = list1.Count;
            for (int index = 0; index < count; ++index)
            {
                Person person1 = people[list1[index]];
                Person person2 = people[list2[index]];
                peoplePairList.Add(new PeoplePair()
                {
                    Person1 = person1,
                    Person2 = person2
                });
            }
            if (list1.Count != list2.Count)
                stragler = people[list2[list2.Count - 1]];
            return peoplePairList;
        }
        private List<PeoplePair> RandomizePeople(
      List<Person> people,
      List<PeoplePair> allArchives,
      ref bool failed,
      ref Person stragler, string archiveFolder)
        {
            List<PeoplePair> peoplePairList = this.CreatePeoplePairList(people, ref stragler);
            string straglerName = "";
            if (stragler != null)
                straglerName = stragler.Name;
            List<Person> source = new List<Person>();
            source.AddRange((IEnumerable<Person>)allArchives.Where<PeoplePair>((Func<PeoplePair, bool>)(x => x.Person1.Name == "")).Select<PeoplePair, Person>((Func<PeoplePair, Person>)(x => x.Person2)).ToList<Person>());
            source.AddRange((IEnumerable<Person>)allArchives.Where<PeoplePair>((Func<PeoplePair, bool>)(x => x.Person2.Name == "")).Select<PeoplePair, Person>((Func<PeoplePair, Person>)(x => x.Person1)).ToList<Person>());
            bool flag1 = true;
            int num1 = 0;
            while (flag1)
            {
                bool flag2 = false;
                foreach (PeoplePair peoplePair in peoplePairList)
                {
                    PeoplePair pp = peoplePair;
                    flag2 = allArchives.Any<PeoplePair>((Func<PeoplePair, bool>)(x =>
                    {
                        if (x.Person1.Name == pp.Person1.Name && x.Person2.Name == pp.Person2.Name)
                            return true;
                        return x.Person1.Name == pp.Person2.Name && x.Person2.Name == pp.Person1.Name;
                    }));
                    if (source.Any<Person>((Func<Person, bool>)(x => x.Name == straglerName)))
                        flag2 = true;
                    if (flag2)
                        break;
                }
                if (flag2)
                {
                    flag1 = true;
                    peoplePairList = this.CreatePeoplePairList(people, ref stragler);
                    if (stragler != null)
                        straglerName = stragler.Name;
                    ++num1;
                    if (num1 >= 1000000)
                    {
                        int num2 = (int)MessageBox.Show("Tool has tried to uniquely pair people 1,000,000 times with no luck. Remove the oldest excel file from " + archiveFolder + " and try again.");
                        failed = true;
                        return new List<PeoplePair>();
                    }
                }
                else
                    flag1 = false;
            }
            failed = false;
            return peoplePairList;
        }
        public class PeoplePair
        {
            public Person Person1 { get; set; }

            public Person Person2 { get; set; }
        }

        public class ArchivePeoplePairs
        {
            public List<PeoplePair> PeoplePairs { get; set; }
        }

        public class Person
        {
            public string Name { get; set; }

            public string Company { get; set; }

            public string Title { get; set; }

            public string Email { get; set; }
        }
    }
}
