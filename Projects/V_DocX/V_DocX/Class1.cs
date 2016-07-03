using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenXmlPowerTools;


namespace V_DocX
{
    public class V_Docx
    {

        public class Segment
        {
            private int id;
            private string text;
            private string type;

            public int Id
            {
                get
                {
                    return id;
                }

                set
                {
                    id = value;
                }
            }

            public string Text
            {
                get
                {
                    return text;
                }

                set
                {
                    text = value;
                }
            }

            public string Type
            {
                get
                {
                    return type;
                }

                set
                {
                    type = value;
                }
            }
        }
        List<Segment> lst = new List<Segment>();
        public V_Docx(){ }
        
        public XmlDocument xdoc = new XmlDocument();
        public string path_Docx;
        public bool Open(string pathDocX)//; throw exception với msg = lý do fail
        {
            path_Docx = pathDocX;
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(pathDocX))
                {
                    //tạo đường dẫn tương đối trong thư mục bin.
                    string directoryPath = "./temp"; 
                    //kiểm tra temp chưa tồn tại thì tạo mới
                    if (!System.IO.Directory.Exists(directoryPath))
                        System.IO.Directory.CreateDirectory(directoryPath);
                    //Delele old Folder
                    Directory.Delete(directoryPath, true);
                    archive.ExtractToDirectory(directoryPath);
                }


                return true;
            }
            catch(Exception tt)
            {
                throw tt;
                return false;
            }
            
        }
        public void Save()
        {
            try
            {
                string s = @"./temp";
                if (File.Exists(path_Docx))
                {
                    File.Delete(path_Docx);

                }

                ZipFile.CreateFromDirectory(s, path_Docx);
            }
            catch (Exception tt)
            {
                throw tt;
            }
        }
        public bool SaveAs(string pathToNewFile)
        {
            try
            {
                string s = @"./temp";
                ZipFile.CreateFromDirectory(s,pathToNewFile);
                return true;
            }
            catch (Exception tt)
            {
                throw tt;
                return false;
            }
        }

        public List<Segment> GetAllSegment()
        {
            List<Segment> lst=new List<Segment>();
            int j = 0;
            string directoryPath = @"./temp/word";
            try
            {
                foreach (string s in Directory.GetFiles(directoryPath))
                {
                    XmlDocument x = new XmlDocument();
                    x.Load(Path.GetFullPath(s));
                    XmlNodeList wt = x.GetElementsByTagName("w:t");
                    for (int i = 0; i < wt.Count; i++)
                    {
                        Segment sg = new Segment();
                        sg.Id = i;
                        sg.Text = wt[i].InnerText;
                        sg.Type = Path.GetFileName(s);
                        lst.Add(sg);
                        wt[i].InnerText = j.ToString();
                        j++;
                    }
                    x.Save(Path.GetFullPath(s));

                }
            }
            catch (Exception tt)
            {
                throw tt;
            }
            return lst;
            
        }

        public string ExtractText()
        {
            return null;

        }

        public void UpdateSegment(Segment _segment)
        {
            XmlDocument x = new XmlDocument();
            string directoryPath = @"./temp/word";
            try
            {
                foreach (string s in Directory.GetFiles(directoryPath))
                {
                    if(Path.GetFileName(s).Equals(_segment.Type))
                    {
                        x.Load(Path.GetFullPath(s));
                        XmlNodeList wt = x.GetElementsByTagName("w:t");
                        wt[_segment.Id].InnerText = _segment.Text;
                        x.Save(Path.GetFullPath(s));
                    }
                }
            }
            catch (Exception tt)
            {
                throw tt;
            }
        }

        public void UpdateSegment(List<Segment> _segments)
        {
            foreach(Segment s in _segments)
            {          
                //Update 1 Segment
                UpdateSegment(s);
            }
        }
        public void UpDateSegment(List<Segment> _segment)
        {
            XmlDocument x = new XmlDocument();
            string directoryPath = @"./temp/word";
            try
            {
                foreach(string s in Directory.GetFiles(directoryPath))
                {
                    x.Load(Path.GetFullPath(s));
                    XmlNodeList wt = x.GetElementsByTagName("w:t");
                    for(int i=0;i<wt.Count;i++)
                    {
                        int id = Convert.ToInt32(wt[i].InnerText);
                        wt[i].InnerText = _segment[id].Text;
                    }
                    x.Save(Path.GetFullPath(s));
                }
            }
            catch(Exception tt)
            {
                throw tt;
            }
        }

        public void UpdateSegmentFile(List<Segment> _segment)
        {
            XmlDocument x = new XmlDocument();
            string directoryPath = @"./temp/word";
            try
            {
                foreach (string s in Directory.GetFiles(directoryPath))
                {
                    if (Path.GetFullPath(s).Equals(_segment[0].Type))
                    {
                        x.Load(Path.GetFullPath(s));
                        XmlNodeList wt = x.GetElementsByTagName("w:t");
                        for (int i = 0; i < _segment.Count; i++)
                        {
                            wt[i].InnerText = _segment[i].Text;
                        }
                        x.Save(Path.GetFullPath(s));
                    }
                }
            }
            catch(Exception tt)
            {
                throw tt;
            }


        }

        public static bool check_Space(string st)
        {
            int dem = 0;
            for (int i = 0; i < st.Count(); i++)
            {
                if (st[i] == 32)
                {
                    dem++;
                }
                if (dem >= 1)
                    return false;
            }
            return true;
        }




        //xử lý file XML
        public void CleanUp()
        {
            string directoryPath = @"./temp/word";
            try
            {
                foreach (string s in Directory.GetFiles(directoryPath))
                {
                    ProcessXML(Path.GetFullPath(s));
                }
            }
            catch(Exception tt)
            {
                throw tt;
            }
        }

        private static void ProcessXML(string Path_XmlFile)       //xử lý file Xml: gộp các phần cùng định dạng
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(Path_XmlFile);
            XmlNodeList del_Space = xdoc.GetElementsByTagName("w:t");
            XmlNodeList wp = xdoc.GetElementsByTagName("w:p");
            for (int ip = 0; ip < wp.Count; ip++)
            {
                XmlNodeList wr = wp[ip].ChildNodes;

                for (int ir = 0; ir < wr.Count - 1; ir++)
                {
                    int dem = 0;

                    if (wr[ir].Name == "w:r")
                    {
                        for (int ir2 = ir + 1; ir2 < wr.Count; ir2++)
                        {
                            bool dd = false;
                            XNodeEqualityComparer equalityComparer = new XNodeEqualityComparer();
                            XmlNode n1 = wr[ir].FirstChild;
                            if (n1 != null)
                            {
                                XElement node1 = n1.GetXElement();
                                XmlNode n2 = wr[ir2].FirstChild;
                                if (wr[ir2].Name == "w:r")
                                {
                                    if (n2 != null)
                                    {
                                        XElement node2 = n2.GetXElement();

                                        //có 1 node và đó là node w:t
                                        if (wr[ir].LastChild.Name == "w:t" && wr[ir].ChildNodes.Count == 1)
                                        {
                                            string s = wr[ir2].LastChild.InnerText;
                                            if (wr[ir].LastChild.Name == "w:t")
                                                wr[ir].LastChild.InnerText += s;
                                            if (wr[ir2].LastChild.Name == "w:t")
                                                wr[ir2].LastChild.InnerText = "";
                                            dd = true;
                                            dem++;
                                        }
                                        else
                                        {
                                            //lỗi chi đây má ơi
                                            if (wr[ir2].LastChild.Name == "w:t")
                                            {
                                                if (equalityComparer.Equals(node1, node2))
                                                {
                                                    string s = wr[ir2].LastChild.InnerText;
                                                    if (wr[ir].LastChild.Name == "w:t")
                                                        wr[ir].LastChild.InnerText += s;
                                                    if (wr[ir2].LastChild.Name == "w:t")
                                                        wr[ir2].LastChild.InnerText = "";
                                                    dd = true;
                                                    dem++;
                                                    //wr[ir2].RemoveAll();    
                                                }
                                                if (wr[ir2].LastChild.InnerText == " ")
                                                {
                                                    if (wr[ir].LastChild.Name == "w:t")
                                                        wr[ir].LastChild.InnerText += " ";
                                                    dd = true;
                                                    wr[ir2].RemoveAll();       // xóa node con không chưa text
                                                }
                                            }
                                        }
                                    }
                                }

                                if (wr[ir2].Name.Equals("w:hyperlink"))
                                {
                                    XmlNode temp = hyperLink(wr[ir2]);
                                }
                                if (wr[ir2].Name.Equals("w:sdt"))
                                {
                                    XmlNode temp2 = wSdt(wr[ir2]);
                                }

                                if (wr[ir2].Name == "w:bookmarkStart" || wr[ir2].Name == "w:bookmarkEnd")
                                {
                                    dd = true;
                                }
                                if (wr[ir2].Name != "w:r")
                                {
                                    dd = true;
                                }
                                if (dd == false)
                                    break;
                            }
                        }
                    }
                    if (wr[ir].Name == "w:hyperlink")
                    {
                        XmlNode temp = hyperLink(wr[ir]);
                    }
                }
            }

            //XmlNodeList elemList = xdoc.GetElementsByTagName("w:t");
            //for (int i = 0; i < elemList.Count; i++)
            //{
            //    //MessageBox.Show(elemList[i].InnerText);
            //    string s = elemList[i].InnerXml;
            //    if (check_Space(s))
            //    {
            //        // MessageBox.Show(elemList[i].InnerXml);
            //        XmlNode cha = elemList[i].ParentNode;
            //        cha.RemoveAll();

            //    }
            //}
            xdoc.Save(Path_XmlFile);

        }
        //process XML file

        public static XmlNode hyperLink(XmlNode x)
        {
            XmlNodeList wr = x.ChildNodes; //get all child of HyperLink..

            for (int ir = 0; ir < wr.Count - 1; ir++)
            {
                int dem = 0;

                if (wr[ir].Name == "w:r")
                {
                    for (int ir2 = ir + 1; ir2 < wr.Count; ir2++)
                    {
                        bool dd = false;
                        XNodeEqualityComparer equalityComparer = new XNodeEqualityComparer();
                        XmlNode n1 = wr[ir].FirstChild;
                        //MessageBox.Show(wr[ir2].Name);
                        if (n1 != null)
                        {
                            XElement node1 = n1.GetXElement();
                            XmlNode n2 = wr[ir2].FirstChild;
                            if (wr[ir2].Name == "w:r")
                            {
                                if (n2 != null)
                                {
                                    XElement node2 = n2.GetXElement();

                                    //có 1 node và đó là node w:t
                                    if (wr[ir].LastChild.Name == "w:t" && wr[ir].ChildNodes.Count == 1)
                                    {
                                        string s = wr[ir2].LastChild.InnerText;
                                        if (wr[ir].LastChild.Name == "w:t")
                                            wr[ir].LastChild.InnerText += s;
                                        if (wr[ir2].LastChild.Name == "w:t")
                                            wr[ir2].LastChild.InnerText = "";
                                        dd = true;
                                        dem++;
                                    }
                                    else
                                    {
                                        //lỗi chi đây má ơi
                                        if (wr[ir2].LastChild.Name == "w:t")
                                        {
                                            if (equalityComparer.Equals(node1, node2))
                                            {
                                                string s = wr[ir2].LastChild.InnerText;
                                                if (wr[ir].LastChild.Name == "w:t")
                                                    wr[ir].LastChild.InnerText += s;
                                                if (wr[ir2].LastChild.Name == "w:t")
                                                    wr[ir2].LastChild.InnerText = "";
                                                dd = true;
                                                dem++;
                                                //wr[ir2].RemoveAll();    
                                            }
                                            if (wr[ir2].LastChild.InnerText == " ")
                                            {
                                                if (wr[ir].LastChild.Name == "w:t")
                                                    wr[ir].LastChild.InnerText += " ";
                                                dd = true;
                                                wr[ir2].RemoveAll();       // xóa node con không chưa text
                                            }
                                        }
                                    }
                                }
                            }

                            if (wr[ir2].Name.Equals("w:hyperlink"))
                            {
                                XmlNode temp = hyperLink(wr[ir2]);
                            }

                            if (wr[ir2].Name.Equals("w:sdt"))
                            {
                                //MessageBox.Show("Có sdt");
                                XmlNode temp2 = wSdt(wr[ir2]);
                            }

                            if (wr[ir2].Name == "w:bookmarkStart" || wr[ir2].Name == "w:bookmarkEnd")
                            {
                                dd = true;
                            }
                            if (wr[ir2].Name != "w:r")
                            {
                                dd = true;
                            }
                            if (dd == false)
                                break;
                        }
                    }
                }
                if (wr[ir].Name == "w:hyperlink")
                {
                    XmlNode temp = hyperLink(wr[ir]);
                }
            }
            XElement node = x.GetXElement();
            return null;
        }      //process one node is hyperlink...

        public static XmlNode wSdt(XmlNode wsdt)
        {
            XmlNodeList lstNode = wsdt.ChildNodes;
            foreach (XmlNode x in lstNode)
            {
                if (x.Name.Equals("w:sdtContent"))      //xử lý node chứa phần mở rộng cho Developer.
                {
                    XmlNodeList wr = x.ChildNodes;
                    for (int ir = 0; ir < wr.Count - 1; ir++)
                    {
                        int dem = 0;

                        if (wr[ir].Name == "w:r")
                        {
                            for (int ir2 = ir + 1; ir2 < wr.Count; ir2++)
                            {
                                bool dd = false;
                                XNodeEqualityComparer equalityComparer = new XNodeEqualityComparer();
                                XmlNode n1 = wr[ir].FirstChild;
                                if (n1 != null)
                                {
                                    XElement node1 = n1.GetXElement();
                                    XmlNode n2 = wr[ir2].FirstChild;
                                    if (wr[ir2].Name == "w:r")
                                    {
                                        if (n2 != null)
                                        {
                                            XElement node2 = n2.GetXElement();

                                            //có 1 node và đó là node w:t
                                            if (wr[ir].LastChild.Name == "w:t" && wr[ir].ChildNodes.Count == 1)
                                            {
                                                string s = wr[ir2].LastChild.InnerText;
                                                if (wr[ir].LastChild.Name == "w:t")
                                                    wr[ir].LastChild.InnerText += s;
                                                if (wr[ir2].LastChild.Name == "w:t")
                                                    wr[ir2].LastChild.InnerText = "";
                                                dd = true;
                                                dem++;
                                            }
                                            else
                                            {
                                                if (wr[ir2].LastChild.Name == "w:t")
                                                {
                                                    if (equalityComparer.Equals(node1, node2))
                                                    {
                                                        string s = wr[ir2].LastChild.InnerText;
                                                        if (wr[ir].LastChild.Name == "w:t")
                                                            wr[ir].LastChild.InnerText += s;
                                                        if (wr[ir2].LastChild.Name == "w:t")
                                                            wr[ir2].LastChild.InnerText = "";
                                                        dd = true;
                                                        dem++;
                                                        //wr[ir2].RemoveAll();    
                                                    }
                                                    if (wr[ir2].LastChild.InnerText == " ")
                                                    {
                                                        if (wr[ir].LastChild.Name == "w:t")
                                                            wr[ir].LastChild.InnerText += " ";
                                                        dd = true;
                                                        wr[ir2].RemoveAll();       // xóa node con không chưa text
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (wr[ir2].Name.Equals("w:hyperlink"))
                                    {
                                        XmlNode temp = hyperLink(wr[ir2]);
                                    }

                                    if (wr[ir2].Name.Equals("w:sdt"))
                                    {
                                        XmlNode temp2 = wSdt(wr[ir2]);
                                    }

                                    if (wr[ir2].Name == "w:bookmarkStart" || wr[ir2].Name == "w:bookmarkEnd")
                                    {
                                        dd = true;
                                    }
                                    if (wr[ir2].Name != "w:r")
                                    {
                                        dd = true;
                                    }
                                    if (dd == false)
                                        break;
                                }
                            }
                        }
                    }

                }
            }
            XElement node = wsdt.GetXElement();
            return null;
        }
        
        public void UpdateFontTable()
        {

        }
        
    }
}
