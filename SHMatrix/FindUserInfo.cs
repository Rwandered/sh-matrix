using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHMatrix
{
    class FindUserInfo
    {
        List<string> uInfo = new List<string>();

        public List<string> FindOneColumnInfo(string OneColName)
        {
            uInfo.Clear();
            string dsd = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string domain = // сюда домен из ini файла
           
            string filter = string.Format("(&(ObjectClass={0})(userPrincipalName={1}))", "person", OneColName + "@" + domain);

            string[] properties = new string[] { "fullname" };

            DirectoryEntry adRoot = new DirectoryEntry("LDAP://" + domain, null, null, AuthenticationTypes.Secure);
            DirectorySearcher searcher = new DirectorySearcher(adRoot);
            searcher.SearchScope = SearchScope.Subtree;
            searcher.ReferralChasing = ReferralChasingOption.All;
            searcher.PropertiesToLoad.AddRange(properties);
            searcher.Filter = filter;

            SearchResult result = searcher.FindOne();
            try
            {
                DirectoryEntry directoryEntry = result.GetDirectoryEntry();     
                UserActivity(directoryEntry.Path.ToString());
                directoryEntry.Close();
                
            }
            catch 
            {
      
                //**************************************************
               
                    string filter2 = string.Format("(&(ObjectClass={0})(sAMAccountName={1}))", "person", OneColName);
                    string[] properties2 = new string[] { "fullname" };
                    DirectoryEntry adRoot2 = new DirectoryEntry("LDAP://" + domain, null, null, AuthenticationTypes.Secure);
                    DirectorySearcher searcher2 = new DirectorySearcher(adRoot2);
                    searcher2.SearchScope = SearchScope.Subtree;
                    searcher2.ReferralChasing = ReferralChasingOption.All;
                    searcher2.PropertiesToLoad.AddRange(properties2);
                    searcher2.Filter = filter2;

                    SearchResult result2 = searcher2.FindOne();
                    try
                    {
                        DirectoryEntry directoryEntry = result2.GetDirectoryEntry();
                        UserActivity(directoryEntry.Path.ToString());
                        directoryEntry.Close();
                    }                  
                    catch
                    {
                    try
                    {
                        using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
                        {
                            GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, OneColName);
                            if (group != null)
                            {
                                uInfo.Add("В группу входят пользователи(группы)");
                                int h = 1;
                                foreach (Principal p in group.GetMembers())
                                {
                                    uInfo.Add(h.ToString() + ". " + p.StructuralObjectClass + " " + p.Name);
                                    h++;
                                }
                            }
                        }
                    }
                    catch { } 

                    }
            }

             return uInfo;
        }

        #region Определить когда логинился пользователь
        void UserActivity(string strPath)
        {
                     
            DirectoryEntry uEntry = new DirectoryEntry(strPath);
            try
            {
                uInfo.Add("Ф.И.О. - " + uEntry.Properties["displayName"].Value.ToString());
            }
            catch { }
            try
            {
                uInfo.Add("Логин - " + uEntry.Properties["samAccountName"].Value.ToString());
            }
            catch { }
            try
            {
                uInfo.Add("E-mail - " + uEntry.Properties["mail"].Value.ToString());
            }
            catch { }
            try
            {
                string dsf = uEntry.Properties["mobile"].Value.ToString();
                uInfo.Add("Номер телефона - " + uEntry.Properties["mobile"].Value.ToString());
            }
            catch (Exception ex)
            {
                
            }
            try
            {
                uInfo.Add("Дата создания пользователя - " + uEntry.Properties["whenCreated"].Value.ToString());
            }
            catch { }
            try
            {
                uInfo.Add("Является членом групп: ");
                int i = 1;
                foreach (object mOf in uEntry.Properties["memberOf"])
                {
                    string mOfs = mOf.ToString();
                    string[] mOfP = mOfs.Split(',');
                    foreach (string m in mOfs.Split(','))
                    {
                        if (m.Contains("CN"))
                        {
                            uInfo.Add(i.ToString() + ". " + m);
                            i++;
                        }
                    }
                }
            }
            catch { }
           
            DirectorySearcher mysearcher = new DirectorySearcher(uEntry);
            SearchResultCollection results = mysearcher.FindAll();
            DirectoryEntry de = results[0].GetDirectoryEntry();
            try
            {
                long lastLogonDateAsLong = GetInt64(de, "lastLogon"); 
                DateTime sfq = DateTime.FromFileTimeUtc(lastLogonDateAsLong);
                uInfo.Add("Дата последнего входа в систему - " + sfq);
            }
            catch { }
            try
            {
                long lastLogonDateAsLong2 = GetInt64(de, "lastLogonTimestamp"); 
            }
            catch { }
            try
            {
                DateTime sf3 = DateTime.FromFileTimeUtc(GetInt64(de, "pwdLastSet"));
                uInfo.Add("Дата последней смены пароля - " + sf3);
            }
            catch { }        
           
        }
        #endregion


        #region Определить когда логинился пользователь

        private static Int64 GetInt64(DirectoryEntry entry, string attr)
        {

            DirectorySearcher ds = new DirectorySearcher(
             entry,
             String.Format("({0}=*)", attr),
             new string[] { attr },
             SearchScope.Base
             );

            SearchResult sr = ds.FindOne();

            if (sr != null)
            {
                if (sr.Properties.Contains(attr))
                {
                    return (Int64)sr.Properties[attr][0];
                }
            }
            return -1;
        }

        #endregion
    }



    static class Datav
    {
        /// <summary>
        /// Логин пользователя
        /// </summary>
        public static string logus { get; set; }
        public static string Value2 { get; set; }
        public static string str { get; set; }

        public static bool portOpen { get; set; }
        public static string modemName { get; set; }
        public static string ds { get; set; }


        /// <summary>
        /// Номер телефона (загружается из AD)
        /// </summary>
        public static string ntel { get; set; }

        public static string sDomain { get; set; }
        public static string prt { get; set; }
        public static string pwd { get; set; }
        public static string portname { get; set; }
        public static DateTime sad { get; set; }

        public static bool FlagReturn { get; set; }

        public static string defaultPath { get; set; }
        public static string sDefaultOU { get; set; }
        public static string sDefaultRootOU { get; set; }
        public static string displayName { get; set; }

    }
}
