using MongoDB.Driver;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using SharpKml.Dom;
using SharpKml.Engine;
using SharpKml.Dom.GX;
using System.IO;
using SharpKml.Base;
using System.Diagnostics;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Security;
using System.DirectoryServices.AccountManagement;
using System.Net.NetworkInformation;
using System.Text;

namespace ProjectTerra
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ReadPlacemark();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"C:\Users\" + Environment.UserName + @"\AppData\LocalLow\Google\GoogleEarth\myplaces.kml");        
            KmlFile file = KmlFile.Load(reader);
            Kml _kml = file.Root as Kml;

            if (_kml != null)
            {
                foreach (var pm in _kml.Flatten().OfType<Placemark>())
                {
                    if (pm.Name.Contains("Eif") || pm.Name.Contains("Christ") || pm.Name.Contains("Google"))
                    {
                       
                    }
                    else
                    {
                        richTextBox1.AppendText(pm.Name.ToString() + "\nLongitude: " + pm.CalculateBounds().Center.Longitude.ToString() + "\nLatitude: " + pm.CalculateBounds().Center.Latitude.ToString() + "\n");
                        WritePlacemark(pm);
                    }           
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"C:\Users\" + Environment.UserName + @"\AppData\LocalLow\Google\GoogleEarth\myplaces.kml");
            KmlFile file = KmlFile.Load(reader);
            Kml _kml = file.Root as Kml;

            if (_kml != null)
            {
                var folder = new Folder();
                folder.Id = "f0";
                folder.Name = "New Pins";

                var placemark = new Placemark();
                placemark.Id = "pm0";
                placemark.Name = "New Location";
                folder.AddFeature(placemark);

                var kml = new Kml();
                kml.Feature = folder;

                var serializer = new Serializer();
                serializer.Serialize(kml);

                placemark = new Placemark();
                placemark.Geometry = new Point { Coordinate = new Vector(38, -120) };
                placemark.Name = "New Location";
                placemark.TargetId = "pm0";

                var update = new Update();
                update.AddUpdate(new ChangeCollection() { placemark });

                serializer.Serialize(update);

                var f = KmlFile.Create(_kml, true);
                update.Process(f);


                using (FileStream stream = File.OpenWrite(@"C:\Users\" + Environment.UserName + @"\myplaces.kml"))
                {
                    f.Save(stream);
                    stream.Close();
                }
            }
        }
        static public bool CheckNetwork()
        {
            try
            {
                Ping mPing = new Ping();
                String host = "google.com";
                byte[] buffer = new byte[32];
                int timeout = 1000;
                PingOptions pingOptions = new PingOptions();
                PingReply reply = mPing.Send(host, timeout, buffer, pingOptions);
                return (reply.Status == IPStatus.Success);
            } catch (Exception)
            {
                return false;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            List<Process> processes = new List<Process>();
            foreach (Process p in Process.GetProcesses())
            {
                if (p.ProcessName.Contains("earth"))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }
        private void WritePlacemark(Placemark pm)
        {
            Serializer sr = new Serializer();
            sr.Serialize(pm);
            String placemark = sr.Xml.Normalize();
            
            using (SqlConnection conn = new SqlConnection(@"Data Source=R9012XG4-X1\SQLEXPRESS;Initial Catalog=GoogleEarth;Integrated Security=True"))
            {
                conn.Open();
                try
                {
                    int substring = placemark.IndexOf(">");
                    placemark = placemark.Substring(substring + 1);
                    using (SqlCommand cmd = new SqlCommand("IF NOT EXISTS (SELECT placemark FROM placemark WHERE placemark='" + placemark + "') BEGIN " +
                        "INSERT INTO placemark (empName, location, placemark) VALUES ('" + Environment.UserName.ToString().ToLower() + "@pinnergy.com', '" + GetLocation() + "', '" + placemark + "') END" , conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                } catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void ReadPlacemark()
        {
            try
            {
                FileInfo fi = new FileInfo(@"C:\users\" + Environment.UserName + @"\AppData\LocalLow\Google\GoogleEarth\myplaces.kml");
                fi.CopyTo(@"C:\users\" + Environment.UserName + @"\myplaces_backup.kml");
                fi.Delete();
            }
            catch (Exception ex)
            {

            }
            using (SqlConnection conn = new SqlConnection(@"Data Source=R9012XG4-X1\SQLEXPRESS;Initial Catalog=GoogleEarth;Integrated Security=True"))
            {
                conn.Open();
                try
                {
                    using (SqlCommand cmd = new SqlCommand("SELECT placemark FROM placemark WHERE location='" + GetLocation() + "'", conn))
                    {
                        SqlDataReader reader = cmd.ExecuteReader();

                        Parser parser = new Parser();
                        Serializer sr = new Serializer();
                        Kml kml = new Kml();
                        Placemark placemark = new Placemark();
                        List<String> points = new List<string>();

                        while (reader.Read())
                        {
                            points.Add(reader.GetString(0));
                        }

                        var folder = new Folder();
                        folder.Id = "points";
                        folder.Name = "Placemarks";

                        foreach (String s in points)
                        {             
                                parser.ParseString(s, false);
                                placemark = (Placemark)parser.Root;
                                folder.AddFeature(placemark);
                                
                                //var update = new SharpKml.Dom.Update();
                                //update.AddUpdate(new ChangeCollection() { placemark });
                                //sr.Serialize(update);
                        }
                        kml.Feature = folder;
                        sr.Serialize(kml);
                        var file = KmlFile.Create(kml, false);
                        file.Save(new FileStream(@"C:\Users\" + Environment.UserName + @"\AppData\LocalLow\Google\GoogleEarth\myplaces.kml", FileMode.Create));
                    }
                } catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private String GetLocation()
        {
            using (var context = new PrincipalContext(ContextType.Domain))
            {
                using (var user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, Environment.UserName))
                {
                    int startIndex = user.DistinguishedName.IndexOf("OU=", 1) + 3;
                    int endIndex = user.DistinguishedName.IndexOf(",", startIndex);
                    var group = user.DistinguishedName.Substring((startIndex), (endIndex - startIndex));
                    return group;
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            
        }
    }
}
