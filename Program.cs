using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using MailBee;
using MailBee.Mime;
using MailBee.Outlook;
using System.Security.Cryptography;

namespace TomesPst2Mime
{
    class Program
    {
        //TomesPst2Mime 
        //A tool using the MailBee .NET objects library to extract mime text from PST files.
        //This build is designed to be a part of the TOMES project and has fixed strings that require its use in a Docker container.
        //This build assumes the exe is running in a mono docker container and the pst files are in the Tomes structure (i.e., /home/tomes/data/pst)
        //Version: 0.0.1
        //Changelog: 
        //=========: * Does basic processing by removing folders that have Delete in the name.
        //=========: * Assigns the EML name as the ItemId from the PST's MAPI.
        //=========: * Normalizes directory paths as lowercase with space's converted to underscores.
        //=========: * Trys to sanitize the crazy stuff people named their Email folders (e.g. inbox\.con, inbox\re:Important). 
        //=========:    * Bad path characters are replaced by spaces
        static void Main(string[] args)
        {
            //   Main(): Executable takes 2 arguments seperated by spaces.
            //   Arg[0]: A pst filename found in the docker mounted volume /home/tomes/data/pst  
            //   Arg[1]: Account Name.
            //=========: The name of the account. 
            //=========:    EX: Gov_Perdue_Correspondence_2017
            //=========:
            //=========:    This will become the top level directory in the docker mounted volume /home/tomes/data/mime_emails 
            //=========:    EX: /home/tomes/data/mime_emails/Gov_Perdue_Correspondence_2017
            //Instructions as of 2018-02-05 on how to implement the license key are found here.  
            //License Key can be found in the Teamwork TOMES project file as a message

            byte[] s = Convert.FromBase64String("***LICENSE KEY WOULD GO HERE***");
            string newS = Encoding.ASCII.GetString(s);
            MailBee.Global.LicenseKey = newS;
            string account = args[0];
            string account_name = args[1];
            PstExtractor extract = new PstExtractor(account, account_name);
            extract.Extract();
        }
    }

    class PstExtractor
    {
        private string pst_file;
        private string mbox_folder = @"/home/tomes/data/mime_emails";
        private string pst_folder = @"/home/tomes/data/pst";
        private PstReader pstRead;
        private PstFolderCollection pstFolders;
        private string collection_root;
        private string sep = @"/";
        private char sep_char = @"/".ToCharArray()[0];
        private IList<string> bad_paths = new List<string>();
        private Dictionary<long, string> good_paths = new Dictionary<long, string>();
        private Dictionary<long, string> path_map = new Dictionary<long, string>();
        private string account_top_name;
        private string account_top_path;
        private string account_name;
        private FileStream folder_map;


        public PstExtractor(string pstFile, string account_name)
        {
            pst_file = Path.GetFullPath(pst_folder + sep + pstFile);
            pstRead = new PstReader(pst_file);
            pstFolders = pstRead.GetPstRootFolders(true);
            collection_root = pstFolders[0].Name.ToLower().Replace(" ", "_");
            this.account_name = account_name;
            mbox_folder += sep + account_name;
            Directory.CreateDirectory(mbox_folder);
            
        }

        public void Extract()
        {
            Regex r = new Regex("Delete.+", RegexOptions.IgnoreCase);
            this.account_top_name = this.pstFolders[0].Name.ToLower().Replace(" ", "_");
            this.account_top_path = $"{this.mbox_folder}" + sep + $"{this.account_top_name}";
            Directory.CreateDirectory(this.account_top_path);
            foreach (PstFolder folder in this.pstFolders)
            {
                string testing_path;
                testing_path = $"{this.mbox_folder}" + sep + $"{folder.Name}";

                MatchCollection mc = r.Matches(testing_path);

                if (mc.Count > 0)
                {
                    continue;
                }

                string path = folder.PstID.ToString();
                this.WriteToFile(path, folder.Name.ToLower());
                string full_path = $"{this.account_top_path}" + sep + $"{path}";
                Directory.CreateDirectory(full_path);
                good_paths[folder.PstID] = full_path;

            }
                
            foreach (PstFolder folder in this.pstFolders) {
                MailMessage inboxMsg = null;
                string folder_path;
                try
                {
                    folder_path = this.good_paths[folder.PstID];
                } catch (Exception)
                {
                    // this would only be an excluded directory
                    continue;
                }
                
                try
                {
                    foreach (PstItem item in folder.Items)
                    {
                        if ((string)item.PstFields["MessageClass"] != "IPM.Note") continue;
                        inboxMsg = item.GetAsMailMessage();
                        string full_path = Path.GetFullPath(folder_path + sep + $"{item.PstID}.eml");
                        Console.WriteLine($"Writing: {full_path}");
                        inboxMsg.SaveMessage(full_path);

                    }
                } catch (ArgumentOutOfRangeException)
                {
                    //
                    continue;
                } catch (System.IndexOutOfRangeException)
                {
                    //
                    continue;
                } 
                
            }
        }


        public void WriteToFile(string id, string path)
        {   
            
            string full_path = Path.Combine($"{this.mbox_folder}", "folder_map.tsv");
            
            if (!File.Exists(full_path))
            {
                Console.WriteLine($"Creating: {full_path}");
                FileStream fs = File.Create(full_path);
                fs.Close();
            }
            
            StreamWriter fh = new StreamWriter(full_path, true);
            fh.WriteLine($"{id}\t{path}");
            fh.Close();
            
        }
    }
}
