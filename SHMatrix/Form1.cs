using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace SHMatrix
{
    public partial class Form1 : Form
    {       
        List<string> ls = new List<string>();
        List<string> ls2 = new List<string>();
        List<string> st1 = new List<string>();
        List<string> st2 = new List<string>();
        List<string> st3 = new List<string>();
        List<string> st4 = new List<string>();
        FindUserInfo FUI = new FindUserInfo();

        List<string> listofRowsMat = new List<string>();
        List<string> listofRowsLMat = new List<string>();

        List<List<string>> GlobalListRowsMat = new List<List<string>>(); 
        List<List<string>> GlobalListRowsLMat = new List<List<string>>();

        Dictionary<string, string> InfoD = new Dictionary<string, string>(25);

        List<string> listRightUserAnaFolder = new List<string>();

        List<string> listRightUserAnaFolder2 = new List<string>();

        List<string> listRightUserAnaFolder3 = new List<string>();


        List<DirectoryInfo> searchdirectory = new List<DirectoryInfo>();


        List<string> ListColNameInfo = new List<string>();


        List<string> ListToInfo = new List<string>();


        List<string> List__name__Cells = new List<string>();

        List<string> List__header__Col = new List<string>();
        

        int i;
        int lsCount;
        List<string> TempList = new List<string>(); 
        List<List<string>> GlobalList = new List<List<string>>(); 
        List<string> uniq;

        System.Timers.Timer t = new System.Timers.Timer();


        List<int> Index_of_Row_No = new List<int>();

        List<int> Index_of_Col_No = new List<int>();


        int indexRow;
        int indexCol;

        int h = 0;
        public Form1()
        {
            InitializeComponent();
            ToolTip tooltip = new ToolTip();
            tooltip.SetToolTip(button4, "Сравнить матрицы");
            Data.FLAG = false;
            Data.compareTrue = false;

            textBox1.Visible = false;

            InfoD.Add("AppendData", "Право добавлять данные в конец файла");
            InfoD.Add("ChangePermissions", "Право на изменение правил безопасности и аудита, связанных с файлом или папкой");
            InfoD.Add("CreateDirectories", "Право на создание папки");
            InfoD.Add("CreateFiles", "Право на создание файла");
            InfoD.Add("Delete", "Право на удаление папки или файла");
            InfoD.Add("DeleteSubdirectoriesAndFiles", "Право на удаление папки и всех файлов в ней");
            InfoD.Add("ExecuteFile", "Право на запуск файла приложения");
            InfoD.Add("FullControl", "Право на полный контроль над папкой или файлом, а также на изменение правил управления доступом и аудита. Это значение представляет право выполнять над файлом любые операции и является объединением всех входящих в перечисление прав");
            InfoD.Add("ListDirectory", "Право на чтение содержимого каталога");
            InfoD.Add("Modify", "Право на чтение, запись, получение содержимого папки, удаление папок и файлов, а также на запуск файлов приложений. Это право включает права ReadAndExecute, Write и Delete");
            InfoD.Add("Read", "Право открывать и копировать папки и файлы с разрешением только для чтения.Это право включает в себя права ReadData, ReadExtendedAttributes, ReadAttributes и ReadPermissions");
            InfoD.Add("ReadAndExecute", "Право открывать и копировать и файлы с разрешением только для чтения, а также запускать файлы приложений.Это право включает права Read и ExecuteFile");
            InfoD.Add("ReadAttributes", "Право открывать и копировать атрибуты файловой системы для папки или файла.Например, это значение дает право на просмотр даты создания или изменения файла.Оно не включает в себя право на чтение данных, дополнительных атрибутов файловой системы или правил доступа и аудита");
            InfoD.Add("ReadData", "Право открывать и копировать файл или папку.Оно не включает в себя право на чтение атрибутов файловой системы, дополнительных атрибутов файловой системы или правил доступа и аудита");
            InfoD.Add("ReadExtendedAttributes", "Право открывать и копировать дополнительные атрибуты файловой системы для папки или файла.Например, оно позволяет просматривать сведения об авторе или содержимом.Оно не включает в себя право на чтение данных, атрибутов файловой системы или правил доступа и аудита");
            InfoD.Add("ReadPermissions", "Право открывать и копировать правила доступа и аудита для папки или файла. Оно не включает в себя право на чтение данных, атрибутов файловой системы и дополнительных атрибутов файловой системы");
            InfoD.Add("Synchronize", "Указывает, может ли приложение ждать синхронизации дескриптора файла с завершением операции ввода - вывода");
            InfoD.Add("TakeOwnership", "Право менять владельца файла или папки. Следует иметь в виду, что владельцы ресурса имеют полный доступ к этому ресурсу");
            InfoD.Add("Traverse", "Право на получение списка содержимого папки и на запуск содержащихся в этой папке приложений");
            InfoD.Add("Write", "Право создавать папки и файлы, а также добавлять данные в файлы и удалять данные из файлов.Это право включает в себя права WriteData, AppendData, WriteExtendedAttributes и WriteAttributes");
            InfoD.Add("WriteAttributes", "Право открытия и записи атрибутов файловой системы для папки или файла. Оно не включает в себя право на запись данных, дополнительных атрибутов файловой системы или правил доступа и аудита");
            InfoD.Add("WriteData", "Право открытия и записи для файла или папки. Оно не включает в себя право на открытие и запись атрибутов файловой системы, дополнительных атрибутов файловой системы или правил доступа и аудита");
            InfoD.Add("WriteExtendedAttributes", "Право открытия и записи дополнительных атрибутов файловой системы для папки или файла.Оно не включает в себя право на запись данных, атрибутов файловой системы или правил доступа и аудита");
           

            pictureBox1.Visible = false;
            label2.Visible = false;

            //*******Properties of menu*********
            openTSMI.Enabled = false;
            saveTSMI.Enabled = false;
            saveAsTSMI.Enabled = false;
            printTSB.Enabled = false;
            printTSMI.Enabled = false;
            tsB1.Enabled = false;
           
            matCompareTSB.Enabled = false;
            matCompareTSMI .Enabled = false;
            button4.Visible = false;
            saveTSB.Enabled = false;
            compareTSB.Enabled = false;
            compareTSMI.Enabled = false;
            clearTSB.Enabled = false;
            clearTSMI.Enabled = false;
            reportTSB.Enabled = false;
            reportTSMI.Enabled = false;

            //*************************


            DataR.top = this.Top;
            DataR.left = this.Left;

            int V = this.Height;
            int SH = this.Width;

            DataR.left = this.Left + SH / 2; 
            DataR.top = this.Top + V/2; 

            FindDir(Data.drive);
            
        }

        #region Функция получения директории
        void FindDir(string drive)
        {
            
            try
            {
                DirectoryInfo dir = new DirectoryInfo(drive);
                DirectoryInfo[] alldir = dir.GetDirectories();
                int alldirCount = alldir.Length;
                if (alldirCount == 0)
                {
                    pictureBox1.Visible = true;
                    label2.Visible = true;
                    
                }
                else
                {
                    WindowsIdentity wi = WindowsIdentity.GetCurrent();
                    BuildMatrix(alldir);
                }
            }
            catch (Exception w)
            {
                if (w.Message.ToString().Contains("Отказано в доступе по пути"))
                {
                    DialogResult result = MessageBox.Show(w.Message, "Ошибка!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Stop,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }
            }
        }
        #endregion


        #region Функция получения директорий и всех поддиректорий
      void FindAllDir(object drive)
        {
            
            try
            {
                DirectoryInfo dir = new DirectoryInfo((string)drive);
                DirectoryInfo[] alldir = dir.GetDirectories();
                for (int i = 0; i < alldir.Length; i++)
                {
                    this.Invoke(new System.Threading.ThreadStart(delegate
                    {
                        searchdirectory.Add(alldir[i]);
                        FindAllDir(alldir[i].FullName);
                        }));
                }
            }
            catch (Exception w)
            { }           
        }
        #endregion
  

        #region Функция построения матрицы
        void BuildMatrix(DirectoryInfo[] alldir)
        {
            try
            {
                System.Threading.Thread thread5 = new System.Threading.Thread(waiting); 
                thread5.Start();

                matrixDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                matrixDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                var col1 = new DataGridViewColumn();
                col1.HeaderText = "Folders/Users/Group";

                //ширина колонки
                col1.ReadOnly = true; 
                col1.Name = "col1"; 
                col1.Frozen = true; 
                col1.CellTemplate = new DataGridViewButtonCell(); 
                matrixDGV.Columns.Add(col1);

                var col2 = new DataGridViewColumn();
                col2.ReadOnly = true; 
                col2.Frozen = true; 
                col2.Name = "col2"; 
                col2.CellTemplate = new DataGridViewTextBoxCell(); 

                try
                {
                    foreach (DirectoryInfo name in alldir)
                    {
                        ls.Add(name.ToString());
                    }
                }
                catch (Exception e)
                { }

                
                foreach (string foldername in ls)
                {
                    
                    matrixDGV.Columns.Add(col2.ToString(), foldername);
                }
                //*****************************************************
                lsCount = ls.Count;

                try
                {
                    foreach (DirectoryInfo name in alldir)
                    {
                        try
                        {

                            DirectorySecurity securityDescriptor = name.GetAccessControl(AccessControlSections.Access);

                            AuthorizationRuleCollection rules = securityDescriptor.GetAccessRules(true, true, typeof(NTAccount));

                            foreach (AuthorizationRule rule in rules)
                            {
                                var fileRule = rule as FileSystemAccessRule;

                                var sf = fileRule.FileSystemRights;

                                string username = fileRule.IdentityReference.Value.ToString();
                                int c = username.LastIndexOf(@"\");
                                if (c == -1)
                                {
                                    ls2.Add(fileRule.IdentityReference.Value.ToString());
                                    
                                    continue;
                                }
                                else
                                {
                                    string usd = username.Remove(0, c + 1);
                                   
                                    ls2.Add(usd);
                                }

                            }
                        }
                        catch (Exception e)
                        { }
                    }

                }
                catch { }

                uniq = ls2.Distinct().ToList(); 


                int ls2Count = uniq.Count;
                for (int i = 0; i < ls2Count; ++i)
                {
                   
                    matrixDGV.Rows.Add();
                }

                i = 0;
                foreach (string IDname in uniq)
                {

                    matrixDGV[0, i].Value = IDname;
                    if (i == ls2Count)
                    {
                        break;
                    }
                    else
                    {
                        i++;
                    }

                }
                foreach (DataGridViewColumn column in matrixDGV.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                saveTSMI.Enabled = true;
                saveAsTSMI.Enabled = true;
                printTSB.Enabled = true;
                printTSMI.Enabled = true;
               
                saveTSB.Enabled = true;
                
                clearTSB.Enabled = true;
                clearTSMI.Enabled = true;

                
                thread5.Abort(); 

            }
            catch { }
        }


        #endregion


        #region Функция построения таблицы по данным из файла основная
        void BuildLoadMatrixOne(string[] matCol, string[] mat)
        {
            try
            {
                System.Threading.Thread thread = new System.Threading.Thread(waiting); //создаем поток, в котором будет открыта 2-я форма
                thread.Start();


                matrixDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                matrixDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                var col1 = new DataGridViewColumn();
                col1.HeaderText = matCol[0];
                var matColL = matCol.ToList();
                matColL.RemoveAt(0);
                string matRowWithCol;
                string[] matRow;
                var matL = mat.ToList();
                matL.RemoveAt(0);
                int matLCout = matL.Count;
               
                col1.ReadOnly = true; 
                col1.Name = "col1"; 
                col1.Frozen = true; 
                col1.CellTemplate = new DataGridViewButtonCell(); 
                matrixDGV.Columns.Add(col1);

                var col2 = new DataGridViewColumn();
                col2.ReadOnly = true; 
                col2.Frozen = true; //флаг, что данная колонка всегда отображается на своем месте
                col2.Name = "col2"; //текстовое имя колонки, его можно использовать вместо обращений по индексу
                col2.CellTemplate = new DataGridViewTextBoxCell(); //тип нашей колонки


                
                foreach (string foldername in matColL)
                {
                   matrixDGV.Columns.Add(col2.ToString(), foldername);
                }
                for (int n = 0; n < matLCout; n++)
                {
                    matRowWithCol = matL[n];
                    matRow = matRowWithCol.Split(';');
                    matrixDGV.Rows.Add(matRow);
                }

                foreach (DataGridViewColumn column in matrixDGV.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                CellsHeadersEmptyOne();

                thread.Abort();
            }
            catch { }
        }
        #endregion


        #region Функция построения таблицы по данным из файла для сравнения
        void BuildLoadMatrixTwo(string[] matCol, string[] mat)
        {
            try
            {
                System.Threading.Thread thread1 = new System.Threading.Thread(waiting); //создаем поток, в котором будет открыта 2-я форма
                thread1.Start();

                loadmatrixDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                loadmatrixDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                var col1LM = new DataGridViewColumn();
                col1LM.HeaderText = matCol[0];
                var matColL = matCol.ToList();
                matColL.RemoveAt(0);
                string matRowWithCol;
                string[] matRow;
                var matL = mat.ToList();
                matL.RemoveAt(0);
                int matLCout = matL.Count;
                //ширина колонки
                col1LM.ReadOnly = true; //значение в этой колонке нельзя править
                col1LM.Name = "col1LM"; //текстовое имя колонки, его можно использовать вместо обращений по индексу
                col1LM.Frozen = true; //флаг, что данная колонка всегда отображается на своем месте
                col1LM.CellTemplate = new DataGridViewButtonCell(); //тип нашей колонки
                loadmatrixDGV.Columns.Add(col1LM);

                var col2LM = new DataGridViewColumn();
                col2LM.ReadOnly = true; //значение в этой колонке нельзя править
                col2LM.Frozen = true; //флаг, что данная колонка всегда отображается на своем месте
                col2LM.Name = "col2LM"; //текстовое имя колонки, его можно использовать вместо обращений по индексу
                col2LM.CellTemplate = new DataGridViewTextBoxCell(); //тип нашей колонки


                // Создадим столбцы матрицы по имени папок
                foreach (string foldername in matColL)
                {
                    //col1.HeaderText = foldername;

                    loadmatrixDGV.Columns.Add(col2LM.ToString(), foldername);
                }
                for (int n = 0; n < matLCout; n++)
                {
                    matRowWithCol = matL[n];
                    matRow = matRowWithCol.Split(';');
                    loadmatrixDGV.Rows.Add(matRow);
                }


                foreach (DataGridViewColumn column in loadmatrixDGV.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                CellsHeadersEmptyTwo();

                thread1.Abort();
            }
            catch { }
        }
        #endregion


        #region Обработчик клика по заголовку
        private void matrixDGV_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            /* try
             {*/
            //Data.drive = Path.Combine(Data.drive, 
            textBox1.Visible = false;
            if (e.Button == MouseButtons.Left)
            {

                if (checkBox1.Checked == true)
                {
                    MessageBox.Show("В режиме сравнения невозможно построить ACL для вложенных каталогов!");
                }
                else
                {
                    string HeaderName = matrixDGV.Columns[e.ColumnIndex].HeaderText;
                    if (HeaderName != "Folders/Users/Group")
                    {
                        Data.drive = Path.Combine(Data.drive, HeaderName);

                        DialogResult result = MessageBox.Show("Построить матрицу доступа для папки: " + HeaderName, "Внимание!",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question,
                               MessageBoxDefaultButton.Button1,

                               MessageBoxOptions.DefaultDesktopOnly);

                        if (result == DialogResult.Yes)
                        {
                            matrixDGV.Columns.Clear();
                            uniq.Clear();
                            ls.Clear();
                            ls2.Clear();
                            TempList.Clear();
                            GlobalList.Clear();
                            FindDir(Data.drive);
                            //return;
                        }
                    }
                }

                /* }
                 catch (Exception n)
                 {

                 }  */
            }

           
        }
        #endregion


        #region Обработчик клика по ячейке
        private void matrixDGV_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
            try
            {
                string sds = matrixDGV.CurrentCell.Value.ToString();


                if (e.ColumnIndex == 0)
                {

                    if (e.RowIndex < 0)
                    {

                    }
                    else
                    {
                        textBox1.Visible = false;
                        System.Threading.Thread thread7 = new System.Threading.Thread(waiting); //создаем поток, в котором будет открыта 2-я форма
                        thread7.Start();
                        textBox1.Clear();
                        int indexRow = matrixDGV.CurrentRow.Index;
                        string TextValue = matrixDGV[0, indexRow].Value.ToString();
                        int CoutLs = FUI.FindOneColumnInfo(TextValue).Count;
                        if (CoutLs == 0)
                        {
                            textBox1.Text = "Локальная группа или пользователь";
                        }
                        else
                        {
                            foreach (string listEl in FUI.FindOneColumnInfo(TextValue))
                            {
                                textBox1.Text += listEl + Environment.NewLine;
                            }
                        }
                        
                        textBox1.Visible = true;
                        thread7.Abort();
                    }
                }
                else
                {
                    if (e.RowIndex < 0)
                    {

                    }
                    else
                    {

                       
                        string userName = matrixDGV[0, e.RowIndex].Value.ToString();
                        string nameDir = Path.Combine(Data.drive, matrixDGV.Columns[e.ColumnIndex].HeaderText);
                        
                        var mousePosition = Cursor.Position;

                        DataR.AbugX = mousePosition.X;
                        DataR.AbugY = mousePosition.Y;

                        try
                        {
                            SetACL sACL = new SetACL();
                            sACL.ShowDialog();
                        }
                        catch (Exception sdfsd) { }

                    }
                }

               
            }
            catch { }
        }
        #endregion



        #region Задаем права доступа н папку
        static public bool SetFullControl(string nameDir, string userName)
        {                                   
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(nameDir);
                DirectorySecurity ds = dirInfo.GetAccessControl(AccessControlSections.Access);

                AuthorizationRuleCollection rules = ds.GetAccessRules(true, true, typeof(NTAccount));

                foreach (AuthorizationRule rule in rules)
                {

                   List<string> TempList = new List<string>();
                    var fileRule = rule as FileSystemAccessRule;

                    TempList.Add(fileRule.AccessControlType.ToString());
                    TempList.Add(fileRule.FileSystemRights.ToString());
                    TempList.Add(fileRule.IdentityReference.ToString());
                }


                ds.AddAccessRule(new FileSystemAccessRule(userName,
             FileSystemRights.Write,
             InheritanceFlags.ContainerInherit,
           PropagationFlags.None,
             AccessControlType.Allow));

                dirInfo.SetAccessControl(ds);


                ds.AddAccessRule(new FileSystemAccessRule(userName,
             FileSystemRights.WriteExtendedAttributes,
             InheritanceFlags.ContainerInherit,
             PropagationFlags.None,
             AccessControlType.Allow));
                dirInfo.SetAccessControl(ds);


                ds.AddAccessRule(new FileSystemAccessRule(userName,
           FileSystemRights.Traverse,
           InheritanceFlags.ContainerInherit,
           PropagationFlags.None,
           AccessControlType.Allow));
                dirInfo.SetAccessControl(ds);


                ds.AddAccessRule(new FileSystemAccessRule(userName,
         FileSystemRights.ChangePermissions,
         InheritanceFlags.ContainerInherit,
         PropagationFlags.None,
         AccessControlType.Allow));
                dirInfo.SetAccessControl(ds);

                ds.AddAccessRule(new FileSystemAccessRule(userName,
        FileSystemRights.ReadPermissions,
        InheritanceFlags.ContainerInherit,
       PropagationFlags.None,
        AccessControlType.Allow));
                dirInfo.SetAccessControl(ds);

              
                dirInfo.SetAccessControl(ds);
                return true;
            }
            catch(Exception evf)
            { return false; }
        }

        #endregion


        #region Функция заполнения матрицы
        void FindAccesList()
        {
            try
            {
                System.Threading.Thread thread2 = new System.Threading.Thread(waiting); //создаем поток, в котором будет открыта 2-я форма
                thread2.Start();

                DirectoryInfo dir = new DirectoryInfo(Data.drive);
                DirectoryInfo[] alldir = dir.GetDirectories();
                WindowsIdentity wi = WindowsIdentity.GetCurrent();

                try
                {
                    foreach (DirectoryInfo Col_name in alldir)
                    {
                        try
                        {
                            DirectorySecurity securityDescriptor = Col_name.GetAccessControl(AccessControlSections.Access);

                            AuthorizationRuleCollection rules = securityDescriptor.GetAccessRules(true, true, typeof(NTAccount));

                            foreach (AuthorizationRule rule in rules)
                            {

                                TempList = new List<string>();
                                var fileRule = rule as FileSystemAccessRule;

                                TempList.Add(fileRule.AccessControlType.ToString());
                                TempList.Add(fileRule.FileSystemRights.ToString());
                                TempList.Add(fileRule.IdentityReference.ToString());

                                GlobalList.Add(TempList); // добавили в общий список элемент (список с атрибутами для одного пользователя)                                                                       

                            }
                            foreach (List<string> TempListR in GlobalList)
                            {
                                Data.Access_type = TempListR[0];
                                Data.Rights = TempListR[1];
                                Data.Ident = TempListR[2];
                                if (Data.Access_type.Equals("Allow"))
                                {
                                    Data.Access_type = "Р: ";
                                }
                                if (Data.Access_type.Equals("Deny"))
                                {
                                    Data.Access_type = "З: ";
                                }
                                if (Data.Rights.Equals("268435456"))
                                {
                                    Data.Rights = "FullControl";
                                }
                                if (Data.Rights.Contains("ReadAndExecute") | Data.Rights.Equals("-1610612736"))
                                {
                                    Data.Rights = "ReadAndExecute";
                                }
                                if (Data.Rights.Contains("Modify") | Data.Rights.Equals("-536805376"))
                                {
                                    Data.Rights = "Modify";
                                }


                                int c = Data.Ident.LastIndexOf(@"\");
                                if (c == -1)
                                {
                                    Data.Ident = TempListR[2];
                                    continue;
                                }
                                else
                                {
                                    Data.Ident = Data.Ident.Remove(0, c + 1);

                                }


                                foreach (string IdentName in uniq)
                                {
                                    if (Data.Ident.Equals(IdentName))
                                    {
                                        indexRow = FindIndexRow(IdentName);
                                        indexCol = FindIndexCol(Col_name.Name/*, indexRow*/);
                                        if (Data.Access_type == " " || Data.Rights == " ")
                                        {
                                            matrixDGV[indexCol, indexRow].Value = "-";
                                            break;
                                        }
                                        else
                                        {
                                            matrixDGV[indexCol, indexRow].Value = Data.Access_type + Data.Rights;
                                            break;
                                        }
                                    }
                                             
                                }
                            }
                            GlobalList.Clear();

                        }
                        catch (Exception e) // ловим ошибки при получении АСL
                        {
                            indexCol = FindIndexCol(Col_name.Name/*, indexRow*/);

                            matrixDGV.Columns[indexCol].DefaultCellStyle.BackColor = System.Drawing.Color.Black;
                            matrixDGV.Columns[indexCol].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                            for (int a = 0; a < uniq.Count; a++)
                            {
                                matrixDGV[indexCol, a].Value = "Error: Нет прав на просмотр ACL";
                            }
                        }

                    }

                }
                catch (Exception e) { }


                CellsEmpty(); // 
                uniq.Clear();
                ls.Clear();
                ls2.Clear();
                TempList.Clear();
                GlobalList.Clear();

                thread2.Abort(); //закрываем поток
            }
            catch { }
        }
        #endregion
        

        #region Функция проверки ячейки на пустоту
        void CellsEmpty()
        {
            int R = matrixDGV.RowCount;
            int K = matrixDGV.ColumnCount;

            for (int i = 0; i < R; i++)
            {
                string stroka = string.Empty;

                for (int j = 0; j < K; j++)
                {
                    if (matrixDGV.Rows[i].Cells[j].Value == null)
                    {
                                                       

                        matrixDGV.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.DarkOrchid;
                        matrixDGV.Rows[i].Cells[j].Style.ForeColor = System.Drawing.Color.White;
                        matrixDGV.Rows[i].Cells[j].Value = "Доступ закрыт";
                    }

                }
            }
        }
        #endregion


        #region Функция проверки ячейки  и заголовка подгруженной матрицы на пустоту для первой матрицы
        void CellsHeadersEmptyOne()
        {
            string noD = "Доступ закрыт";
            string noACL = "Error: Нет прав на просмотр ACL";
            int R = matrixDGV.RowCount;
            int K = matrixDGV.ColumnCount;

            for (int i = 0; i < R; i++)
            {
                string stroka = string.Empty;

                for (int j = 0; j < K; j++)
                {

                    string TextVarlue = matrixDGV.Columns[j].HeaderText.ToString();

                    if (TextVarlue == "\t")
                    {
                        Data.IndexColLoadgM = j;

                    }

                    if (matrixDGV.Rows[i].Cells[j].Value.Equals(noD))
                    {
                        
                        matrixDGV.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.DarkOrchid;
                        matrixDGV.Rows[i].Cells[j].Style.ForeColor = System.Drawing.Color.White;
                       
                    }
                    if (matrixDGV.Rows[i].Cells[j].Value.Equals(noACL))
                    {
                 
                        matrixDGV.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Black;
                        matrixDGV.Rows[i].Cells[j].Style.ForeColor = System.Drawing.Color.White;
                       
                    }

                }
            }

            matrixDGV.Columns.RemoveAt(Data.IndexColLoadgM);// удалим пустой столбец
        }
        #endregion
        

        #region Функция проверки ячейки  и заголовка подгруженной матрицы на пустоту для второй матрицы
        void CellsHeadersEmptyTwo()
        {
            string noD = "Доступ закрыт";
            string noACL = "Error: Нет прав на просмотр ACL";
            int R = loadmatrixDGV.RowCount;
            int K = loadmatrixDGV.ColumnCount;

            for (int i = 0; i < R; i++)
            {
                string stroka = string.Empty;
                

                for (int j = 0; j < K; j++)
                {                   
                    string TextVarlue = loadmatrixDGV.Columns[j].HeaderText.ToString();

                    if (TextVarlue == "\t")
                    {
                        Data.IndexColLoadgM = j;
                            
                    }
                    
                    if (loadmatrixDGV.Rows[i].Cells[j].Value.Equals(noD))
                    {
                                                          

                        loadmatrixDGV.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.DarkOrchid;
                        loadmatrixDGV.Rows[i].Cells[j].Style.ForeColor = System.Drawing.Color.White;
                       
                    }
                    if (loadmatrixDGV.Rows[i].Cells[j].Value.Equals(noACL))
                    {
                        loadmatrixDGV.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Black;
                        loadmatrixDGV.Rows[i].Cells[j].Style.ForeColor = System.Drawing.Color.White;
                        
                    }

                }
            }

            loadmatrixDGV.Columns.RemoveAt(Data.IndexColLoadgM);// удалим пустой столбец
        }
        #endregion


        #region Поиск индекса строки
        int FindIndexRow(string ValuuRowCol)
        {
            
            for (int s = 0; s < uniq.Count+1; s++)
            {
                string TextVarlue = matrixDGV[0, s].Value.ToString();
                if (TextVarlue == ValuuRowCol)
                {
                    indexRow = s;
                    break;
                }
            }
            return indexRow;
        }
        #endregion
        

        #region Поиск индекса столбца
        int FindIndexCol(string ValuuRowCol)
        {

            for (int s = 0; s < ls.Count+1; s++)
            {
                string TextVarlue = matrixDGV.Columns[s].HeaderText.ToString();
                if (TextVarlue == ValuuRowCol)
                {
                   
                        indexCol = s;
                        break;
                }
            }
            return indexCol;
        }
        #endregion
        

        #region Обработчики кнопок
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            if (label4.Visible == true)
            {
                label4.Visible = false;
            }
            textBox1.Visible = false;
            if (pictureBox1.Visible == false)
            {

                if (Data.drive == null)
                {
                    DialogResult result = MessageBox.Show("Не задан путь до каталога!", "Ошибка!",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Error,
                     MessageBoxDefaultButton.Button1,

                     MessageBoxOptions.DefaultDesktopOnly);
                }

                if (ls.Count == 0 || ls2.Count == 0)
                {

                }
                else
                {
                    try
                    {
                        FindAccesList();
                    }
                    catch { }
                }
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            label2.Visible = false;
            textBox1.Visible = false;

            if (folderDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = folderDialog1.SelectedPath;
                Data.drive = textBox4.Text;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            textBox1.Visible = false;
            if (label4.Visible == true)
            {
                label4.Visible = false;
            }
            pictureBox1.Visible = false;
            label2.Visible = false;

            if (Data.drive == null)
            {
                DialogResult result = MessageBox.Show("Не задан путь до каталога!", "Ошибка!",
                 MessageBoxButtons.OK,
                 MessageBoxIcon.Error,
                 MessageBoxDefaultButton.Button1,

                 MessageBoxOptions.DefaultDesktopOnly);
            }
            matrixDGV.Columns.Clear();

            try
            {
                uniq.Clear();
                ls.Clear();
                ls2.Clear();
                TempList.Clear();
                GlobalList.Clear();
            }
            catch { }

            FindDir(Data.drive);

        }


      
        private void Form1_Move(object sender, EventArgs e)
        {
            int V = this.Height;
            int SH = this.Width;
            DataR.top = this.Top;
            DataR.left = this.Left;
            DataR.left = this.Left + SH / 2; // координата центра X
            DataR.top = this.Top + V / 2; // координата центра по У
           
        }


        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            try
            {
                string fileCSV = "";
                saveFileDialog1.Filter = "Таблицы(*.csv)|*.csv|All files(*.*)|*.*";
                saveFileDialog1.ShowDialog();
                // string name = textBox1.Text;
                for (int f = 0; f < matrixDGV.ColumnCount; f++)
                {
                    fileCSV += (matrixDGV.Columns[f].HeaderText + ";");

                }
                fileCSV += "\t\n"; //тут была загвоздка
                for (int i = 0; i < matrixDGV.RowCount; i++)
                {

                    for (int j = 0; j < matrixDGV.ColumnCount; j++)
                    {

                        fileCSV += (matrixDGV[j, i].Value).ToString() + ";";
                    }

                    fileCSV += "\t\n";
                }
                StreamWriter wr = new StreamWriter(saveFileDialog1.FileName, false, Encoding.GetEncoding("windows-1251"));
                wr.Write(fileCSV);
                wr.Close();
                label4.Visible = true;
            }
            catch
            {
                DialogResult result = MessageBox.Show("Не удалось сохранить файл!", "Ошибка!",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,

                MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            try
            {
            pictureBox1.Visible = false;
            label2.Visible = false;

                if (matrixDGV.Columns.Count != 0) // если построена матрица
                {
                    DialogResult result = MessageBox.Show("Матрица прав доступа уже октрыта!\nВы уверены, что хотите открыть новую матрицу и заменить текущую?", "Внимание!",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1,

                    MessageBoxOptions.DefaultDesktopOnly);
                    if (result == DialogResult.Yes)
                    {
                        matrixDGV.Columns.Clear();
                        openFileDialog1.Filter = "Таблицы(*.csv)|*.csv";
                        openFileDialog1.FileName = " ";
                        openFileDialog1.ShowDialog();
                        string[] mat = File.ReadAllLines(openFileDialog1.FileName, Encoding.GetEncoding("windows-1251"));

                        int matCout = mat.Length; // число строк матрицы
                        string[] matCol = mat[0].Split(';');
                        int matColCout = matCol.Length - 1; // число столбцов матрицы

                        BuildLoadMatrixOne(matCol, mat);
                        if (loadmatrixDGV.ColumnCount != 0)
                        {
                            bool CompareMat = RowOneColOneCompare();

                            if (CompareMat == false)
                            {
                                DialogResult result1 = MessageBox.Show("Невозможно сравнить данные матрицы!\n", "Ошибка!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,

                                MessageBoxOptions.DefaultDesktopOnly);
                                try
                                {
                                    matrixDGV.Columns.Clear();
                                }
                                catch
                                { }

                            }
                            else
                            {
                                button4.Visible = true;
                            }
                        }
                    }
                }
                else
                {
                    openFileDialog1.Filter = "Таблицы(*.csv)|*.csv|All files(*.*)|*.*";
                    openFileDialog1.FileName = " ";
                    openFileDialog1.ShowDialog();
                    string[] mat = File.ReadAllLines(openFileDialog1.FileName, Encoding.GetEncoding("windows-1251"));

                    int matCout = mat.Length; // число строк матрицы
                    string[] matCol = mat[0].Split(';');
                    int matColCout = matCol.Length - 1; // число столбцов матрицы

                    BuildLoadMatrixOne(matCol, mat);
                    if (loadmatrixDGV.ColumnCount != 0)
                    {
                        bool CompareMat = RowOneColOneCompare();

                        if (CompareMat == false)
                        {
                            DialogResult result1 = MessageBox.Show("Невозможно сравнить данные матрицы!\n", "Ошибка!",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,

                            MessageBoxOptions.DefaultDesktopOnly);
                            try
                            {
                                matrixDGV.Columns.Clear();
                            }
                            catch
                            { }

                        }
                        else
                        {
                            button4.Visible = true;
                        }
                    }
                }
            }
            catch (Exception b)
            {
                DialogResult result = MessageBox.Show("Не удалось открыть файл!\n" + b.Message, "Ошибка!",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,

                MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            if (matrixDGV.Columns.Count != 0)
            {
                try
                {

                    pictureBox1.Visible = false;
                    label2.Visible = false;


                    if (loadmatrixDGV.Columns.Count != 0) // если построена матрица
                    {
                        DialogResult result = MessageBox.Show("Матрица прав доступа уже октрыта!\nВы уверены, что хотите открыть новую матрицу и заменить текущую?", "Внимание!",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1,

                        MessageBoxOptions.DefaultDesktopOnly);
                        if (result == DialogResult.Yes)
                        {
                            loadmatrixDGV.Columns.Clear();
                            openFileDialog1.Filter = "Таблицы(*.csv)|*.csv|All files(*.*)|*.*";
                            openFileDialog1.FileName = " ";
                            openFileDialog1.ShowDialog();
                            string[] mat = File.ReadAllLines(openFileDialog1.FileName, Encoding.GetEncoding("windows-1251"));

                            int matCout = mat.Length; // число строк матрицы
                            string[] matCol = mat[0].Split(';');
                            int matColCout = matCol.Length - 1; // число столбцов матрицы

                            BuildLoadMatrixTwo(matCol, mat);
                           bool CompareMat = RowOneColOneCompare();

                            if (CompareMat == false)
                            {
                                DialogResult result1 = MessageBox.Show("Невозможно сравнить данные матрицы!\n", "Ошибка!",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1,

                                MessageBoxOptions.DefaultDesktopOnly);
                                try
                                {
                                    loadmatrixDGV.Columns.Clear();
                                }
                                catch
                                { }

                            }
                            else
                            {
                                button4.Visible = true;
                            }
                        }
                    }
                    else
                    {
                        openFileDialog1.Filter = "Таблицы(*.csv)|*.csv";
                        openFileDialog1.FileName = " ";
                        openFileDialog1.ShowDialog();
                        string[] mat = File.ReadAllLines(openFileDialog1.FileName, Encoding.GetEncoding("windows-1251"));

                        int matCout = mat.Length; // число строк матрицы
                        string[] matCol = mat[0].Split(';');
                        int matColCout = matCol.Length - 1; // число столбцов матрицы

                        BuildLoadMatrixTwo(matCol, mat);
                        bool CompareMat = RowOneColOneCompare();
                        if (CompareMat == false)
                        {
                            DialogResult result1 = MessageBox.Show("Невозможно сравнить данные матрицы!\n", "Ошибка!",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,

                            MessageBoxOptions.DefaultDesktopOnly);
                            loadmatrixDGV.Columns.Clear();

                        }
                        else
                        {
                            button4.Visible = true;
                        }
                    }


                }
                catch (Exception b)
                {
                    DialogResult result = MessageBox.Show("Не удалось открыть файл!\n" + b.Message, "Ошибка!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,

                    MessageBoxOptions.DefaultDesktopOnly);
                }
            }
            else // пока нет основной матрицы нельзя добавить вторую
            {
                    DialogResult result = MessageBox.Show("Ошибка в добавлении матрицы для сравнения!\nНет основной матрицы.", "Ошибка!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,

                    MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private void очиститьПоляToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            try
            {
                

                DialogResult result = MessageBox.Show("Вы действительно хотите очистить актиные поля?", "Внимание!",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Question,
                          MessageBoxDefaultButton.Button1,

                          MessageBoxOptions.DefaultDesktopOnly);

                if (result == DialogResult.Yes)
                {
                    matrixDGV.Columns.Clear();
                    loadmatrixDGV.Columns.Clear();

                    if (checkBox1.Checked == false)
                    {
                        openTSMI.Enabled = false;
                        saveTSMI.Enabled = false;
                        saveAsTSMI.Enabled = false;
                        printTSB.Enabled = false;
                        printTSMI.Enabled = false;
                        tsB1.Enabled = false;
                       
                        saveTSB.Enabled = false;
                        compareTSB.Enabled = false;
                        compareTSMI.Enabled = false;
                        clearTSB.Enabled = false;
                        clearTSMI.Enabled = false;
                        reportTSB.Enabled = false;
                        reportTSMI.Enabled = false;
                    }
                    else
                    {
                        saveTSMI.Enabled = true;
                        saveAsTSMI.Enabled = true;
                        printTSB.Enabled = true;
                        printTSMI.Enabled = true;
                        
                        saveTSB.Enabled = true;
                        
                        clearTSB.Enabled = true;
                        clearTSMI.Enabled = true;

                    }
                }



            }
            catch (Exception m)
            { }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            try
            {
                if (matrixDGV.Columns.Count != 0 && loadmatrixDGV.Columns.Count != 0)
                {
                    checkBox1.Checked = true;
                }
                else
                {

                    TSL.Visible = true;
                    if (checkBox1.Checked == false) // режим формирования матрицы
                    {
                        TSL.Text = "Режим формирования ACL";

                        openTSMI.Enabled = false;
                        saveTSMI.Enabled = false;
                        saveAsTSMI.Enabled = false;
                        printTSB.Enabled = false;
                        printTSMI.Enabled = false;
                        tsB1.Enabled = false;
                        tsB2.Enabled = true;
                        tsB3.Enabled = true;
                        tsB4.Enabled = true;
                        matCompareTSB.Enabled = false;
                        matCompareTSMI.Enabled = false;
                        button4.Visible = false;
                        saveTSB.Enabled = false;
                        compareTSB.Enabled = false;
                        compareTSMI.Enabled = false;
                        clearTSB.Enabled = false;
                        clearTSMI.Enabled = false;
                        reportTSB.Enabled = false;
                        reportTSMI.Enabled = false;
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        textBox4.Enabled = true;

                        if (matrixDGV.Columns.Count != 0)
                        {
                            saveTSMI.Enabled = true;
                            saveAsTSMI.Enabled = true;
                            printTSB.Enabled = true;
                            printTSMI.Enabled = true;
                            
                            saveTSB.Enabled = true;
                           
                            clearTSB.Enabled = true;
                            clearTSMI.Enabled = true;

                        }

                    }
                    if (checkBox1.Checked == true) // режим сравнения
                    {

                        TSL.Text = "Режим сравнения ACL";
                        openTSMI.Enabled = true;
                        saveTSMI.Enabled = true;
                        saveAsTSMI.Enabled = true;
                        printTSB.Enabled = true;
                        printTSMI.Enabled = true;
                        tsB1.Enabled = true;
                        tsB2.Enabled = false;
                        tsB3.Enabled = false;
                        tsB4.Enabled = false;
                        saveTSB.Enabled = true;
                        compareTSB.Enabled = true;
                        compareTSMI.Enabled = true;
                        clearTSB.Enabled = true;
                        clearTSMI.Enabled = true;

                        reportTSB.Enabled = true;
                        reportTSMI.Enabled = true;
                        button1.Enabled = false;
                        button2.Enabled = false;
                        button3.Enabled = false;
                        textBox4.Enabled = false;

                        matCompareTSB.Enabled = true;
                        matCompareTSMI.Enabled = true;
                        

                        if (loadmatrixDGV.Columns.Count == 0)
                        {
                            button4.Visible = true;
                        }


                    }
                }
            }
            catch
            {

            }
        }



        private void printTSB_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            PrintDocument Document = new PrintDocument();
            Document.DefaultPageSettings.Landscape = true;
            Document.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
            //printDialog1.ShowDialog();
            PrintPreviewDialog dlg = new PrintPreviewDialog();
            dlg.Document = Document;
            dlg.ShowDialog();
            Document.Print();
        }

        #endregion


        #region Функция вызова формы ожидания
        void waiting()
        {
            try
            {
                Progress f2 = new Progress();
                f2.ShowDialog(Owner as Form1);
            }
            catch { }
            
        }


        #endregion

        
       private bool SetupThePrinting()
        {
            PrintDialog MyPrintDialog = new PrintDialog();
            MyPrintDialog.AllowCurrentPage = false;
            MyPrintDialog.AllowPrintToFile = false;
            MyPrintDialog.AllowSelection = false;
            MyPrintDialog.AllowSomePages = false;
            MyPrintDialog.PrintToFile = false;
            MyPrintDialog.ShowHelp = false;
            MyPrintDialog.ShowNetwork = false;

            if (MyPrintDialog.ShowDialog() != DialogResult.OK)
                return false;

           printDocument1.DocumentName = this.Text;
            printDocument1.PrinterSettings =
                                MyPrintDialog.PrinterSettings;
            printDocument1.DefaultPageSettings =
            MyPrintDialog.PrinterSettings.DefaultPageSettings;
            printDocument1.DefaultPageSettings.Landscape = true;
            printDocument1.DefaultPageSettings.Margins =
                             new Margins(30, 30, 30, 30);

            DataGridViewPrinter MyDataGridViewPrinter = new DataGridViewPrinter(matrixDGV,
             printDocument1, true, true, this.Text, new Font("Tahoma", 8,
                FontStyle.Bold, GraphicsUnit.Point), Color.Black, true);

            return true;
        }

      
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            Graphics g = e.Graphics;
            int x = 0;
            int y = 20;
            int cell_height = 0;

            int colCount = matrixDGV.ColumnCount;
            int rowCount = matrixDGV.RowCount - 1;

            Font font = new Font("Palatino Linotype", 9, FontStyle.Bold, GraphicsUnit.Point);

            int[] widthC = new int[colCount];

            int current_col = 0;
            int current_row = 0;

            while (current_col < colCount)
            {
                if (g.MeasureString(matrixDGV.Columns[current_col].HeaderText.ToString(), font).Width > widthC[current_col])
                {
                    widthC[current_col] = (int)g.MeasureString(matrixDGV.Columns[current_col].HeaderText.ToString(), font).Width;
                }
                current_col++;
            }

            while (current_row < rowCount)
            {
                while (current_col < colCount)
                {
                    if (g.MeasureString(matrixDGV[current_col, current_row].Value.ToString(), font).Width > widthC[current_col])
                    {
                        widthC[current_col] = (int)g.MeasureString(matrixDGV[current_col, current_row].Value.ToString(), font).Width;
                    }
                    current_col++;
                }
                current_col = 0;
                current_row++;
            }

            current_col = 0;
            current_row = 0;

            string value = "";

            int width = widthC[current_col];
            int height = matrixDGV[current_col, current_row].Size.Height;

            Rectangle cell_border;
            SolidBrush brush = new SolidBrush(Color.Black);


            while (current_col < colCount)
            {
                width = widthC[current_col];
                cell_height = matrixDGV[current_col, current_row].Size.Height;
                cell_border = new Rectangle(x, y, width, height);
                value = matrixDGV.Columns[current_col].HeaderText.ToString();
                g.DrawRectangle(new Pen(Color.Black), cell_border);
                g.DrawString(value, font, brush, x, y);
                x += widthC[current_col];
                current_col++;
            }
            while (current_row < rowCount)
            {
                while (current_col < colCount)
                {
                    width = widthC[current_col];
                    cell_height = matrixDGV[current_col, current_row].Size.Height;
                    cell_border = new Rectangle(x, y, width, height);
                    value = matrixDGV[current_col, current_row].Value.ToString();
                    g.DrawRectangle(new Pen(Color.Black), cell_border);
                    g.DrawString(value, font, brush, x, y);
                    x += widthC[current_col];
                    current_col++;
                }
                current_col = 0;
                current_row++;
                x = 0;
                y += cell_height;
            }
        }

        private void matrixDGV_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                if (e.RowIndex < 0)
                {

                }
                else
                {
                    textBox1.Visible = false;
                    ListToInfo.Clear();
                    try
                    {

                        if (e.Button == MouseButtons.Right)
                        {

                            System.Threading.Thread thread9 = new System.Threading.Thread(waiting); //создаем поток, в котором будет открыта 2-я форма
                            thread9.Start();

                            DataGridViewRow clickedRow = (sender as DataGridView).Rows[e.RowIndex];
                            if (!clickedRow.Selected)
                                matrixDGV.CurrentCell = clickedRow.Cells[e.ColumnIndex];

                            var mousePosition = Cursor.Position;

                            int indexRow = matrixDGV.CurrentRow.Index;
                            string TextValue = matrixDGV[e.ColumnIndex, indexRow].Value.ToString();

                            int CoutLs = FUI.FindOneColumnInfo(TextValue).Count;
                            if (CoutLs == 0)
                            {
                                ListToInfo.Add("Локальная группа или пользователь");
                            }
                            else
                            {
                                foreach (string listEl in FUI.FindOneColumnInfo(TextValue))
                                {
                                    ListToInfo.Add(listEl);
                                }
                            }

                            DataR.AbugX = mousePosition.X;
                            DataR.AbugY = mousePosition.Y;

                            try
                            {
                                AbUG ab = new AbUG(ListToInfo);
                                thread9.Abort();
                                ab.ShowDialog();
                            }
                            catch (Exception sdfsd) { }


                        }
                    }
                    catch (Exception efd)
                    { }
                }
            }
        }
              #region Функция сравнения матриц
        private void button4_Click(object sender, EventArgs e)
        {
            List__name__Cells.Clear();
            List__header__Col.Clear();


            Data.Fl__Index = 0;
            textBox1.Visible = false;
            if (label4.Visible == true)
            {
                label4.Visible = false;
            }
            listRightUserAnaFolder.Clear();
            if (loadmatrixDGV.Columns.Count == 0)
            {
                    DialogResult result = MessageBox.Show("Ошибка в сравнении матриц!\nОтсутствуют элементы для сравнения.", "Ошибка!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,

                    MessageBoxOptions.DefaultDesktopOnly);
            }
            else
            {
                FindDifRowAndCol();
                     

                for (int y = 0; y < matrixDGV.RowCount; y++) // тут мы проходим по матрице 1 и значения заголовком делаем стандартными
                {
                                                            
                        matrixDGV.Rows[y].Cells[0].Style.BackColor = System.Drawing.SystemColors.Control;
                        matrixDGV.Rows[y].Cells[0].Style.ForeColor = System.Drawing.SystemColors.WindowText;                                     
                }

                for (int y = 0; y < loadmatrixDGV.RowCount; y++) // тут мы проходим по матрице 2 и значения заголовком делаем стандартными
                {                
                  loadmatrixDGV.Rows[y].Cells[0].Style.BackColor = System.Drawing.SystemColors.Control;
                  loadmatrixDGV.Rows[y].Cells[0].Style.ForeColor = System.Drawing.SystemColors.WindowText;
                }

                for (int matIr = 0; matIr < matrixDGV.RowCount; matIr++)
                {
                    for (int matIc = 0; matIc < matrixDGV.ColumnCount; matIc++)
                    {
                        if (matrixDGV.Rows[matIr].Cells[matIc].Style.BackColor == Color.Red)
                        {
                            listRightUserAnaFolder.Add(matrixDGV.Rows[matIr].Cells[0].Value.ToString() + ";" +
                               matrixDGV.Columns[matrixDGV.Rows[matIr].Cells[matIc].ColumnIndex].HeaderText.ToString()+ ";" +
                               matrixDGV.Rows[matIr].Cells[matIc].Value.ToString());
                            Data.FLAG = true;

                        }

                        if (matrixDGV.Rows[matIr].DefaultCellStyle.BackColor == Color.Blue)
                        {
                            listRightUserAnaFolder2.Add(matrixDGV.Rows[matIr].Cells[0].Value.ToString());
                            listRightUserAnaFolder2 = listRightUserAnaFolder2.Distinct().ToList();
                            Data.FLAG = true;

                        }
                        if (matrixDGV.Columns[matIc].DefaultCellStyle.BackColor == Color.Blue)
                        {
                            listRightUserAnaFolder3.Add(matrixDGV.Columns[matIc].HeaderText); 
                            listRightUserAnaFolder3 = listRightUserAnaFolder3.Distinct().ToList();
                            Data.FLAG = true;

                        }                       
                    }
                }

                for (int matIr = 0; matIr < loadmatrixDGV.RowCount; matIr++)
                {
                    for (int matIc = 0; matIc < loadmatrixDGV.ColumnCount; matIc++)
                    {                       

                        if (loadmatrixDGV.Rows[matIr].DefaultCellStyle.BackColor == Color.Blue)
                        {
                            listRightUserAnaFolder2.Add(loadmatrixDGV.Rows[matIr].Cells[0].Value.ToString());
                              
                            listRightUserAnaFolder2 = listRightUserAnaFolder2.Distinct().ToList();
                            Data.FLAG = true;

                        }
                        if (loadmatrixDGV.Columns[matIc].DefaultCellStyle.BackColor == Color.Blue)
                        {
                            listRightUserAnaFolder3.Add(loadmatrixDGV.Columns[matIc].HeaderText); 
                               
                            listRightUserAnaFolder3 = listRightUserAnaFolder3.Distinct().ToList();
                            Data.FLAG = true;

                        }
                    }
                }
                
                if (Data.FLAG == false) // одинаковые матрицы
                {
                    Data.compareTrue = true;
                    DialogResult result = MessageBox.Show("Сравнение  прав доступа ACL завершено!\nПрава доступа ACL не изменены!\nПостроить отчет по сравнению прав доступа ACL?", "Готово!",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                    if (result == DialogResult.Yes)
                    {
                        reportTSB.PerformClick();
                    }
                }
                else
                {
                    Data.compareTrue = true;
                    DialogResult result = MessageBox.Show("Сравнение прав доступа ACL завершено!\nПрава доступа ACL изменены!\nПостроить отчет по сравнению прав доступа ACL?", "Готово!",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                    if (result == DialogResult.Yes)
                    {
                        reportTSB.PerformClick();
                    }
                }
                
            }

            listofRowsLMat.Clear();
            GlobalListRowsLMat.Clear();
            GlobalListRowsMat.Clear();
            listofRowsMat.Clear();

            

        }
        #endregion


        #region Function to find different cells
        void FindDifRowAndCol() // разница между столбцами и стркоами
        {
            if (matrixDGV.RowCount == loadmatrixDGV.RowCount) // число строк равно
            {
                if (matrixDGV.ColumnCount == loadmatrixDGV.ColumnCount) // равно ли число столбцов
                {
                    FindDifCellsWhenCellEqual(); // функция когда все равно

                }
                else // число столбов не равно / значит либо у первой матрицы больше столбцов, либо у второй
                {
                    if (matrixDGV.ColumnCount > loadmatrixDGV.ColumnCount) // число столбцов первой матрицы больше
                    {
                        FindDifCellsWhenCellEqualButColNoEq(matrixDGV, loadmatrixDGV);// функция когда стркои равны но столбов больше у первой матрицы
                    }
                    else // число столбцов второй матрицы больше
                    {
                        FindDifCellsWhenCellEqualButColNoEq(loadmatrixDGV, matrixDGV);// функция когда стркои равны но столбов больше у второй матрицы
                    }
                }

            }
            else // число строк не равно
            {
                if (matrixDGV.ColumnCount == loadmatrixDGV.ColumnCount) // равно ли число столбцов
                {
                    if (matrixDGV.RowCount > loadmatrixDGV.RowCount)// функция когда строки не равны, но столбцы равны
                    {
                        FindDifCellsWhenCellNoEqualButColEq(matrixDGV, loadmatrixDGV);
                    }
                    else
                    {
                        FindDifCellsWhenCellNoEqualButColEq(loadmatrixDGV, matrixDGV);
                    }                                   

                }
                else // число столбов не равно / значит либо у первой матрицы больше столбцов, либо у второй
                {
                   
                    if (matrixDGV.ColumnCount > loadmatrixDGV.ColumnCount && matrixDGV.RowCount > loadmatrixDGV.RowCount) // число столбцов первой матрицы больше
                    {
                        
                        FindDifCellsWhenCellNoEqualAndColNiEq(matrixDGV, loadmatrixDGV);
                       
                   }
                    else // число столбцов второй матрицы больше 
                    {
                        
                      FindDifCellsWhenCellNoEqualAndColNiEq(loadmatrixDGV, matrixDGV);                                             
                                         
                    }
                }

            }
           
        }
       

        #region Поиск строки по имени первой ячейки
        int FindFirstCell(string nameCell, DataGridView matrix) // принимает 2 параметра 1 это имя ячейки, 2 - 2а матрица
        {
            int indexRowNow = -1; // индекс нужной строки
            for (int i = 0; i < matrix.RowCount; i++)
            {
                if (matrix.Rows[i].Cells[0].Value.ToString() == nameCell)
                {
                    indexRowNow = i;
                    List__name__Cells.Add(nameCell);
                }
            }

            return indexRowNow;
        }
        #endregion
        

        #region Поиск столбца по заголовку
        int FindColByHeader(string nameHeader, DataGridView matrix) // принимает 2 параметра 1 это имя ячейки, 2 - 2а матрица
        {
           
            int indexColNow = -1; // индекс нужной строки
            for (int i = 0; i < matrix.ColumnCount; i++)
            {
                if (matrix.Columns[i].HeaderText == nameHeader)
                {
                    indexColNow = i;
                    List__header__Col.Add(nameHeader);
                }
            }

            return indexColNow;
        }
        #endregion

        
        #region Функция когда столбцы и стркои одинакоое количество
        void FindDifCellsWhenCellEqual() // функция, когда все равно
        {
            int RowIn = 0; // индекс строкиж
            for (int i = 0; i < matrixDGV.RowCount; i++)
            {

                RowIn = FindFirstCell(matrixDGV.Rows[i].Cells[0].Value.ToString(), loadmatrixDGV); //найдем строку во 2 матрице
                
                for (int j = 0; j < matrixDGV.ColumnCount; j++)
                {
                    if (RowIn == -1)
                    {
                        matrixDGV.Rows[i].DefaultCellStyle.BackColor = Color.Blue;
                        matrixDGV.Rows[i].DefaultCellStyle.ForeColor = Color.White;
                        
                    }
                    else
                    {
                        if (!matrixDGV.Rows[i].Cells[j].Value.ToString().Equals(loadmatrixDGV.Rows[RowIn].Cells[j].Value.ToString())) // не равный
                        {
                            matrixDGV.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            matrixDGV.Rows[i].Cells[j].Style.ForeColor = Color.White;
                            loadmatrixDGV.Rows[RowIn].Cells[j].Style.BackColor = Color.Red;
                            loadmatrixDGV.Rows[RowIn].Cells[j].Style.ForeColor = Color.White;
                        }
                    }


                }

            }

            // закрасим во 2 матрицы строки не совпадения



            for (int h = 0; h < loadmatrixDGV.RowCount; h++)
            {
               
                    string ds = loadmatrixDGV.Rows[h].Cells[0].Value.ToString();
                   // string ds2 = List__name__Cells[h];

                    if (List__name__Cells.IndexOf(ds) < 0)
                    {
                        loadmatrixDGV.Rows[h].DefaultCellStyle.BackColor = Color.Blue;
                        loadmatrixDGV.Rows[h].DefaultCellStyle.ForeColor = Color.White;
                        //break;
                    }
                
            }

        }
        #endregion


        #region  функция когда строки не равны, но столбцы равны
        void FindDifCellsWhenCellNoEqualButColEq(DataGridView matrixBig, DataGridView matrixSmole)
        {


            int RowIn = 0; // индекс строкиж
            for (int i = 0; i < matrixBig.RowCount; i++)
            {

                RowIn = FindFirstCell(matrixBig.Rows[i].Cells[0].Value.ToString(), matrixSmole); //найдем строку во 2 матрице

                for (int j = 0; j < matrixBig.ColumnCount; j++)
                {
                    if (RowIn == -1)
                    {
                        matrixBig.Rows[i].DefaultCellStyle.BackColor = Color.Blue;
                        matrixBig.Rows[i].DefaultCellStyle.ForeColor = Color.White;

                    }
                    else
                    {
                        if (!matrixBig.Rows[i].Cells[j].Value.ToString().Equals(matrixSmole.Rows[RowIn].Cells[j].Value.ToString())) // не равный
                        {
                            matrixBig.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            matrixBig.Rows[i].Cells[j].Style.ForeColor = Color.White;
                            matrixSmole.Rows[RowIn].Cells[j].Style.BackColor = Color.Red;
                            matrixSmole.Rows[RowIn].Cells[j].Style.ForeColor = Color.White;
                        }
                    }


                }

            }

            // закрасим во 2 матрицы строки не совпадения
            for (int h = 0; h < matrixSmole.RowCount; h++)
            {

                string ds = matrixSmole.Rows[h].Cells[0].Value.ToString();
                // string ds2 = List__name__Cells[h];

                if (List__name__Cells.IndexOf(ds) < 0)
                {
                    matrixSmole.Rows[h].DefaultCellStyle.BackColor = Color.Blue;
                    matrixSmole.Rows[h].DefaultCellStyle.ForeColor = Color.White;
                    //break;
                }

            }


        }
        #endregion

        #region функция когда стркои равны но столбцов нет
        void FindDifCellsWhenCellEqualButColNoEq(DataGridView matrixBig, DataGridView matrixSmole)
        {
            int ColIn = 0; // индекс столбца
            for (int i = 0; i < matrixBig.ColumnCount; i++)
            {
                //matrixDGV.Columns[i].HeaderText;
                ColIn = FindColByHeader(matrixBig.Columns[i].HeaderText, matrixSmole); //найдем столбец во 2 матрице

                for (int j = 0; j < matrixBig.RowCount; j++)
                {
                    if (ColIn == -1)
                    {
                        matrixBig.Columns[i].DefaultCellStyle.BackColor = Color.Blue;
                        matrixBig.Columns[i].DefaultCellStyle.ForeColor = Color.White;

                    }
                    else
                    {
                        if (!matrixBig.Rows[j].Cells[i].Value.ToString().Equals(matrixSmole.Rows[j].Cells[ColIn].Value.ToString())) // не равный
                        {
                            matrixBig.Rows[j].Cells[ColIn].Style.BackColor = Color.Red;
                            matrixBig.Rows[j].Cells[ColIn].Style.ForeColor = Color.White;
                            matrixSmole.Rows[j].Cells[ColIn].Style.BackColor = Color.Red;
                            matrixSmole.Rows[j].Cells[ColIn].Style.ForeColor = Color.White;
                        }
                    }


                }

            }
            // закрасим во 2 матрицы строки не совпадения
            for (int h = 0; h < matrixSmole.ColumnCount; h++)
            {

                string ds = matrixSmole.Columns[h].HeaderText;
                // string ds2 = List__name__Cells[h];

                if (List__header__Col.IndexOf(ds) < 0)
                {
                    matrixSmole.Columns[h].DefaultCellStyle.BackColor = Color.Blue;
                    matrixSmole.Columns[h].DefaultCellStyle.ForeColor = Color.White;
                    //break;
                }

            }
        }
        #endregion

        #region Функция когда строки разные и столбы разные
        void FindDifCellsWhenCellNoEqualAndColNiEq(DataGridView matrixBig, DataGridView matrixSmole)
        {
            int ColIn = 0; // индекс столбца
            int RowIn = 0; // индекс строки

            for (int i = 0; i < matrixBig.RowCount; i++)
            {
                RowIn = FindFirstCell(matrixBig.Rows[i].Cells[0].Value.ToString(), matrixSmole); //найдем строку во 2 матрице

                for (int j = 0; j < matrixBig.ColumnCount; j++)
                {
                    ColIn = FindColByHeader(matrixBig.Columns[j].HeaderText, matrixSmole); //найдем столбец во 2 матрице

                    if (ColIn == -1 & RowIn == -1 )
                    {
                        matrixBig.Columns[j].DefaultCellStyle.BackColor = Color.Blue;
                        matrixBig.Columns[j].DefaultCellStyle.ForeColor = Color.White;

                        matrixBig.Rows[i].DefaultCellStyle.BackColor = Color.Blue;
                        matrixBig.Rows[i].DefaultCellStyle.ForeColor = Color.White;

                    }
                    else // тут может быть 2 варианта 1) оба > -1 2) RowIn == -1 ColIn > -1 3) RowIn > -1 ColIn == -1
                    {
                        if (ColIn != -1 & RowIn != -1)
                        {
                            if (!matrixBig.Rows[i].Cells[j].Value.ToString().Equals(matrixSmole.Rows[RowIn].Cells[ColIn].Value.ToString()))
                            {
                                matrixBig.Rows[i].Cells[j].Style.BackColor = Color.Red;
                                matrixBig.Rows[i].Cells[j].Style.ForeColor = Color.White;
                                matrixSmole.Rows[RowIn].Cells[ColIn].Style.BackColor = Color.Red;
                                matrixSmole.Rows[RowIn].Cells[ColIn].Style.ForeColor = Color.White;
                            }
                        }

                        if (ColIn != -1 & RowIn == -1)
                        {
                            matrixBig.Rows[i].DefaultCellStyle.BackColor = Color.Blue;
                            matrixBig.Rows[i].DefaultCellStyle.ForeColor = Color.White;
                        }

                        if (ColIn == -1 & RowIn != -1)
                        {
                            matrixBig.Columns[j].DefaultCellStyle.BackColor = Color.Blue;
                            matrixBig.Columns[j].DefaultCellStyle.ForeColor = Color.White;
                        }


                    }
                }
            }
                      }
        #endregion

        #endregion




        #region Функция сравнения первой строки и 1 столбца матрицы
        bool RowOneColOneCompare()
        {
            Index_of_Col_No.Clear();
            Index_of_Row_No.Clear();
            st1.Clear();
            st2.Clear();
            st3.Clear();
            st4.Clear();
            
            string HeaderOfMainMat = matrixDGV.Columns[0].HeaderText.ToString();

            string HeaderOfSecondMat = loadmatrixDGV.Columns[0].HeaderText.ToString();

            if (HeaderOfMainMat != HeaderOfSecondMat) // если заголовки первого столбца разные - такие матрицы не сравниваем вообще, это значит, что она создана не в этой проге
            {
                return false;
            }
            else // если заголовки первого столбца равные 
            {

                int matC = matrixDGV.RowCount; // число строк первой матрицы
                int lmatC = loadmatrixDGV.RowCount; // число строк второй матрицы

                int matR = matrixDGV.ColumnCount; // число столбцов первой матрицы
                int lmatR = loadmatrixDGV.ColumnCount; // число столбцов второй мтарицы

                if (matC == lmatC) // число строк равны
                {
                    if (matR == lmatR) // число столбцов равны
                    {
                        FindElByMatrix(matC, lmatC, matR, lmatR); // заполним списки

                        

                    }
                    else // число строк равно а столбцов нет
                    {
                        FindElByMatrix(matC, lmatC, matR, lmatR);

                        #region Попробую найти индексы столбца
                        if (st3.Count > st4.Count) // если число строк основной матрицы меньше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st3.Count; j++) // в итоге пробегаем по тому списку, который больше
                            {
                                var fl = st4.IndexOf(st3[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Col_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                    Data.Name__of__matrix = 0; //"matrixDGV";
                                                               //+ надо еще какой0то фла о том, какой именно список
                                }
                            }
                        }
                        else // если число строк основной матрицы больше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st4.Count; j++) // идем по элементам списка
                            {
                                var fl = st3.IndexOf(st4[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Col_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                                            //+ надо еще какой0то фла о том, какой именно список
                                    Data.Name__of__matrix = 1; //"loadmatrixDGV";
                                }
                            }

                        }
                        #endregion
                        Data.Switch__matrix__size = 2; // строки одинаковые столбцы нет
                    }
                }
                else // число срок не равно
                {
                    if (matR == lmatR) // проверим равно ли тогда число столбцов
                    {

                        FindElByMatrix(matC, lmatC, matR, lmatR); // заполним списки

                        #region Попробую найти индексы строки
                        if (st1.Count > st2.Count) // если число строк основной матрицы меньше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st1.Count; j++) // в итоге пробегаем по тому списку, который больше
                            {
                                var fl = st2.IndexOf(st1[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Row_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                    Data.Name__of__matrix = 0; //"matrixDGV";
                                                               //+ надо еще какой0то фла о том, какой именно список
                                }
                            }
                        }
                        else // если число строк основной матрицы больше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st2.Count; j++) // идем по элементам списка
                            {
                                var fl = st1.IndexOf(st2[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Row_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                                            //+ надо еще какой0то фла о том, какой именно список
                                    Data.Name__of__matrix = 1; //"loadmatrixDGV";
                                }
                            }

                        }
                        #endregion

                        Data.Switch__matrix__size = 1; // число столбцов равно, а число строк разное
                    }
                    else // число столбцов не равно, а значит и матрицы совсем разные - разное число и строк и столбцов 
                    {
                        FindElByMatrix(matC, lmatC, matR, lmatR); // заполним списки


                        #region Попробую найти индексы строки
                        if (st1.Count > st2.Count) // если число строк основной матрицы меньше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st1.Count; j++) // в итоге пробегаем по тому списку, который больше
                            {
                                var fl = st2.IndexOf(st1[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Row_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                    Data.Name__of__matrix = 0; //"matrixDGV";
                                                               //+ надо еще какой0то фла о том, какой именно список
                                }
                            }
                        }
                        else // если число строк основной матрицы больше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st2.Count; j++) // идем по элементам списка
                            {
                                var fl = st1.IndexOf(st2[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Row_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                                            //+ надо еще какой0то фла о том, какой именно список
                                    Data.Name__of__matrix = 1; //"loadmatrixDGV";
                                }
                            }

                        }
                        #endregion

                        #region Попробую найти индексы столбца
                        if (st3.Count > st4.Count) // если число строк основной матрицы меньше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st3.Count; j++) // в итоге пробегаем по тому списку, который больше
                            {
                                var fl = st4.IndexOf(st3[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Col_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                    Data.Name__of__matrix = 0; //"matrixDGV";
                                                               //+ надо еще какой0то фла о том, какой именно список
                                }
                            }
                        }
                        else // если число строк основной матрицы больше числа строк матрицы сравнения
                        {
                            for (int j = 0; j < st4.Count; j++) // идем по элементам списка
                            {
                                var fl = st3.IndexOf(st4[j]);

                                if (fl < 0) // значит нет такого элемента
                                {
                                    Index_of_Col_No.Add(j); // и индекс этого элемента записываем (запоминаем)
                                                            //+ надо еще какой0то фла о том, какой именно список
                                    Data.Name__of__matrix = 1; //"loadmatrixDGV";
                                }
                            }

                        }
                        #endregion

                        Data.Switch__matrix__size = 3; // м разневе и столбцы и строки
                    }
                }

                return true;
            }
        }
        #endregion



        #region Функция получения списков строк и столбцов матрицы для нахождения разных элементов
        void FindElByMatrix(int matC, int lmatC, int matR, int lmatR)
        {
            for (int i = 0; i < matC; i++)
            {

                st1.Add(matrixDGV.Rows[i].Cells[0].Value.ToString()); // список  имен строк основной матрицы

            }

            for (int j = 0; j < matR; j++)
            {

                st3.Add(matrixDGV.Columns[j].HeaderText.ToString()); // список столбцов основной матрицы

            }


            for (int k = 0; k < lmatC; k++)
            {

                st2.Add(loadmatrixDGV.Rows[k].Cells[0].Value.ToString()); // список имен строк матрицы сравнения
            }

            for (int n = 0; n < lmatR; n++)
            {


                st4.Add(loadmatrixDGV.Columns[n].HeaderText.ToString()); // список столбцов матрицы сравнения

            }
        }
        #endregion



        private void matrixDGV_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if ((e.ColumnIndex != 0 && e.Value != null))
            {
                DataGridViewCell cell = matrixDGV.Rows[e.RowIndex].Cells[e.ColumnIndex];

                string TextValue = matrixDGV[e.ColumnIndex, e.RowIndex].Value.ToString();

                foreach (KeyValuePair<string, string> keyValue in InfoD)
                {
                    if (TextValue.Equals("Р: " + keyValue.Key) | TextValue.Equals("З: " + keyValue.Key))
                    {
                        //MessageBox.Show(keyValue.Key + " - " + keyValue.Value);
                        // tBTMS.Text = keyValue.Key + " - " + keyValue.Value;
                       
                        cell.ToolTipText = keyValue.Key + " - " + keyValue.Value;
                        
                        break;
                    }

                }
            }
        }


        #region Формирование отчета 
        private void reportTSB_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            try
            {
                if (Data.compareTrue == false)
                {
                    DialogResult result = MessageBox.Show("Формирование отчета доступно после сравнения ACL!", "Ошибка!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,

                    MessageBoxOptions.DefaultDesktopOnly);
                }
                else
                {
                    var fileHTML = DataGridtoHTML(matrixDGV, loadmatrixDGV);

                    saveFileDialog1.Filter = "Html(*.html)|*.html";
                    saveFileDialog1.ShowDialog();
                    // string name = textBox1.Text;

                    StreamWriter wr = new StreamWriter(saveFileDialog1.FileName, false, Encoding.GetEncoding("windows-1251"));
                    wr.Write(fileHTML);
                    wr.Close();
                }
            }
            catch
            {
                DialogResult result = MessageBox.Show("Не удалось сохранить файл!", "Ошибка!",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,

                MessageBoxOptions.DefaultDesktopOnly);
            }
        }



        private StringBuilder DataGridtoHTML(DataGridView dg, DataGridView dg2)
        {
            StringBuilder strB = new StringBuilder();
            strB.AppendLine("<!DOCTYPE html>");
            strB.AppendLine("<html><body>");
            strB.AppendLine("<p align = 'center' >" + "<b> Отчет по сравнению прав доступа (ACL) к каталогам</b></p>");

            strB.AppendLine("<ol>");

            strB.AppendLine("<li style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b>&nbsp;Матрица прав доступа ACL </b></font></font></li>");
            strB.AppendLine("<br>");
            strB.AppendLine("<table border='1' cellpadding='0' cellspacing='0'>");
            strB.AppendLine("<tr>");

            for (int i = 0; i < dg.Columns.Count; i++)
            {
                strB.AppendLine("<td align='center' valign='middle'>" +
                               dg.Columns[i].HeaderText + "</td>");
            }

            strB.AppendLine("<tr>");
            for (int i = 0; i < dg.RowCount; i++)
            {
                string rowcolor = dg.Rows[i].DefaultCellStyle.BackColor.Name.ToString();
                rowcolor = rowcolor.ToLower();
                if (rowcolor.Equals("0"))
                {
                    strB.AppendLine("<tr>");
                }
                else
                {
                    strB.AppendLine("<tr style='color:white; background-color: " + rowcolor + "'>");
                }
                foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)
                {
                    string color = dgvc.Style.BackColor.Name.ToString();
                    color = color.ToLower();
                    if (color.Equals("0") | color.Equals("control"))
                    {
                        strB.AppendLine("<td align='center' valign='middle'>" +
                                      dgvc.Value.ToString() + "</td>");
                    }
                    else
                    {
                        strB.AppendLine("<td align='center' style='color:white; background-color: " + color + "' valign='middle'>" +
                                      dgvc.Value.ToString() + "</td>");
                    }
                }
                strB.AppendLine("</tr>");

            }
            strB.AppendLine("</table>");
            strB.AppendLine("<br>");

            strB.AppendLine("<li style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b>&nbsp;Матрица сравнения прав доступа ACL </b></font></li>");
            strB.AppendLine("<br>");

            strB.AppendLine("<" +
              "table border='1' cellpadding='0' cellspacing='0'>");
            strB.AppendLine("<tr>");

            for (int i = 0; i < dg2.Columns.Count; i++)
            {
                strB.AppendLine("<td align='center' valign='middle'>" +
                               dg2.Columns[i].HeaderText + "</td>");
            }

            strB.AppendLine("<tr>");
            for (int i = 0; i < dg2.Rows.Count; i++)
            {

                string rowcolor = dg2.Rows[i].DefaultCellStyle.BackColor.Name.ToString();
                rowcolor = rowcolor.ToLower();
                if (rowcolor.Equals("0"))
                {
                    strB.AppendLine("<tr>");
                }
                else
                {
                    strB.AppendLine("<tr style='color:white; background-color: " + rowcolor + "'>");
                }

                
                foreach (DataGridViewCell dgvc in dg2.Rows[i].Cells)
                {
                    string color = dgvc.Style.BackColor.Name.ToString();
                    
                    color = color.ToLower();
                   
                    
                    if (color.Equals("0") | color.Equals("control"))
                    {
                        strB.AppendLine("<td align='center' valign='middle'>" +
                                      dgvc.Value.ToString() + "</td>");
                    }
                    else
                    {                        
                        strB.AppendLine("<td align='center' style='color:white; background-color: " + color + "' valign='middle'>" +
                                     dgvc.Value.ToString() + "</td>");
                    }
                }
                strB.AppendLine("</tr>");

            }
            strB.AppendLine("</table>");
            strB.AppendLine("<br>");
            if (Data.FLAG == false)
            {
                strB.AppendLine("<li style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b>&nbsp;Права доступа для данных каталогов не изменены!</b></font></font></li>");
                strB.AppendLine("<br>");


                strB.AppendLine("</ol>");
                strB.AppendLine("</body></html>");
            }
            else
            {
                strB.AppendLine("<li style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b>&nbsp;Имеются несогласованные права доступа: </b></font></font></li>");
                strB.AppendLine("<br>");
                strB.AppendLine("</ol>");
                foreach (string abt in listRightUserAnaFolder)
                {
                    string[] abtP = abt.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    
                    strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b><li>&nbsp;Пользователь(группа): "+ abtP[0] + "</b> имеет не согласованные права доступа(<b>"+abtP[2]+"</b>) к каталогу: <b>"+abtP[1]+"</b></font></font></ul>");
                    strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b><li>&nbsp;Сведения о пользователе(группе): </b></font></font></ul>");
                    int CoutLs = FUI.FindOneColumnInfo(abtP[0]).Count;
                    if (CoutLs == 0)
                    {
                        string abtUG = "Локальная группа или пользователь";
                        strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><li>&nbsp;" + abtUG + "</font></font></ul>");
                    }
                    else
                    {
                        foreach (string abtUG in FUI.FindOneColumnInfo(abtP[0]))
                        {
                            strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><li>&nbsp;" + abtUG + "</font></font></ul>");
                        }
                    }
                                       
                    strB.AppendLine("<br>");
                }

                foreach (string abt in listRightUserAnaFolder2)
                {
                    

                    strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b><li>&nbsp;Пользователь(группа): " + abt + "</b> не имел доступа к данной группе каталогов. </font></font></ul>");
                    strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b><li>&nbsp;Сведения о пользователе(группе): </b></font></font></ul>");
                    int CoutLs = FUI.FindOneColumnInfo(abt).Count;
                    if (CoutLs == 0)
                    {
                        string abtUG = "Локальная группа или пользователь";
                        strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><li>&nbsp;" + abtUG + "</font></font></ul>");
                    }
                    else
                    {
                        foreach (string abtUG in FUI.FindOneColumnInfo(abt))
                        {
                            strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><li>&nbsp;" + abtUG + "</font></font></ul>");
                        }
                    }

                    strB.AppendLine("<br>");
                }

                foreach (string abt in listRightUserAnaFolder3)
                {


                    strB.AppendLine("<ul style='text - align: justify; '><font size='3'><font face='Times New Roman, Times, serif'><b><li>&nbsp;В данной группе каталогов создан новый вложенный каталог с именем: </b>" + abt + "</font></font></ul>");
                    

                    strB.AppendLine("<br>");
                }


                strB.AppendLine("</body></html>");
            }

           
            return strB;
        }

        #endregion

        private void helpTSP_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
        }        
    }


    public class Data
    {
        public static string Access_type { get; set; }
        public static string Rights { get; set; }
        public static string Ident { get; set; }
        public static string drive { get; set; }
        public static int IndexColLoadgM { get; set; }
        public static bool FLAG { get; set; }
        public static bool compareTrue { get; set; }
       
        public static int Name__of__matrix { get; set; }
        public static int Switch__matrix__size { get; set; }

        public static int Fl__Index { get; set; }

        public static int ii { get; set; }
    }


   class DataGridViewPrinter
    {
        private static DataGridView TheDataGridView; // The DataGridView Control which will be printed
        private static PrintDocument ThePrintDocument; // The PrintDocument to be used for printing
        private static bool IsCenterOnPage; // Determine if the report will be printed in the Top-Center of the page
        private static bool IsWithTitle; // Determine if the page contain title text
        private static string TheTitleText; // The title text to be printed in each page (if IsWithTitle is set to true)
        private static Font TheTitleFont; // The font to be used with the title text (if IsWithTitle is set to true)
        private static Color TheTitleColor; // The color to be used with the title text (if IsWithTitle is set to true)
        private static bool IsWithPaging; // Determine if paging is used

        static int CurrentRow; // A static parameter that keep track on which Row (in the DataGridView control) that should be printed

        static int PageNumber;

        private static int PageWidth;
        private static int PageHeight;
        private static int LeftMargin;
        private static int TopMargin;
        private static int RightMargin;
        private static int BottomMargin;

        private static float CurrentY; // A parameter that keep track on the y coordinate of the page, so the next object to be printed will start from this y coordinate

        private static float RowHeaderHeight;
        private static List<float> RowsHeight;
        private static List<float> ColumnsWidth;
        private static float TheDataGridViewWidth;

        // Maintain a generic list to hold start/stop points for the column printing
        // This will be used for wrapping in situations where the DataGridView will not fit on a single page
        private static List<int[]> mColumnPoints;
        private static List<float> mColumnPointsWidth;
        private static int mColumnPoint;

        // The class constructor
        public DataGridViewPrinter(DataGridView aDataGridView, PrintDocument aPrintDocument, bool CenterOnPage, bool WithTitle, string aTitleText, Font aTitleFont, Color aTitleColor, bool WithPaging)
        {
            TheDataGridView = aDataGridView;
            ThePrintDocument = aPrintDocument;
            IsCenterOnPage = CenterOnPage;
            IsWithTitle = WithTitle;
            TheTitleText = aTitleText;
            TheTitleFont = aTitleFont;
            TheTitleColor = aTitleColor;
            IsWithPaging = WithPaging;

            PageNumber = 0;

            RowsHeight = new List<float>();
            ColumnsWidth = new List<float>();

            mColumnPoints = new List<int[]>();
            mColumnPointsWidth = new List<float>();

            // Claculating the PageWidth and the PageHeight
            if (!ThePrintDocument.DefaultPageSettings.Landscape)
            {
                PageWidth = ThePrintDocument.DefaultPageSettings.PaperSize.Width;
                PageHeight = ThePrintDocument.DefaultPageSettings.PaperSize.Height;
            }
            else
            {
                PageHeight = ThePrintDocument.DefaultPageSettings.PaperSize.Width;
                PageWidth = ThePrintDocument.DefaultPageSettings.PaperSize.Height;
            }

            // Claculating the page margins
            LeftMargin = ThePrintDocument.DefaultPageSettings.Margins.Left;
            TopMargin = ThePrintDocument.DefaultPageSettings.Margins.Top;
            RightMargin = ThePrintDocument.DefaultPageSettings.Margins.Right;
            BottomMargin = ThePrintDocument.DefaultPageSettings.Margins.Bottom;

            // First, the current row to be printed is the first row in the DataGridView control
            CurrentRow = 0;
        }

        // The function that calculate the height of each row (including the header row), the width of each column (according to the longest text in all its cells including the header cell), and the whole DataGridView width
        private static void Calculate(Graphics g)
        {
            if (PageNumber == 0) // Just calculate once
            {
                SizeF tmpSize = new SizeF();
                Font tmpFont;
                float tmpWidth;

                TheDataGridViewWidth = 0;
                for (int i = 0; i < TheDataGridView.Columns.Count; i++)
                {
                    tmpFont = TheDataGridView.ColumnHeadersDefaultCellStyle.Font;
                    if (tmpFont == null) // If there is no special HeaderFont style, then use the default DataGridView font style
                        tmpFont = TheDataGridView.DefaultCellStyle.Font;

                    tmpSize = g.MeasureString(TheDataGridView.Columns[i].HeaderText, tmpFont);
                    tmpWidth = tmpSize.Width;
                    RowHeaderHeight = tmpSize.Height;

                    for (int j = 0; j < TheDataGridView.Rows.Count; j++)
                    {
                        tmpFont = TheDataGridView.Rows[j].DefaultCellStyle.Font;
                        if (tmpFont == null) // If the there is no special font style of the CurrentRow, then use the default one associated with the DataGridView control
                            tmpFont = TheDataGridView.DefaultCellStyle.Font;

                        tmpSize = g.MeasureString("Anything", tmpFont);
                        RowsHeight.Add(tmpSize.Height);

                        tmpSize = g.MeasureString(TheDataGridView.Rows[j].Cells[i].EditedFormattedValue.ToString(), tmpFont);
                        if (tmpSize.Width > tmpWidth)
                            tmpWidth = tmpSize.Width;
                    }
                    if (TheDataGridView.Columns[i].Visible)
                        TheDataGridViewWidth += tmpWidth;
                    ColumnsWidth.Add(tmpWidth);
                }

                // Define the start/stop column points based on the page width and the DataGridView Width
                // We will use this to determine the columns which are drawn on each page and how wrapping will be handled
                // By default, the wrapping will occurr such that the maximum number of columns for a page will be determine
                int k;

                int mStartPoint = 0;
                for (k = 0; k < TheDataGridView.Columns.Count; k++)
                    if (TheDataGridView.Columns[k].Visible)
                    {
                        mStartPoint = k;
                        break;
                    }

                int mEndPoint = TheDataGridView.Columns.Count;
                for (k = TheDataGridView.Columns.Count - 1; k >= 0; k--)
                    if (TheDataGridView.Columns[k].Visible)
                    {
                        mEndPoint = k + 1;
                        break;
                    }

                float mTempWidth = TheDataGridViewWidth;
                float mTempPrintArea = (float)PageWidth - (float)LeftMargin - (float)RightMargin;

                // We only care about handling where the total datagridview width is bigger then the print area
                if (TheDataGridViewWidth > mTempPrintArea)
                {
                    mTempWidth = 0.0F;
                    for (k = 0; k < TheDataGridView.Columns.Count; k++)
                    {
                        if (TheDataGridView.Columns[k].Visible)
                        {
                            mTempWidth += ColumnsWidth[k];
                            // If the width is bigger than the page area, then define a new column print range
                            if (mTempWidth > mTempPrintArea)
                            {
                                mTempWidth -= ColumnsWidth[k];
                                mColumnPoints.Add(new int[] { mStartPoint, mEndPoint });
                                mColumnPointsWidth.Add(mTempWidth);
                                mStartPoint = k;
                                mTempWidth = ColumnsWidth[k];
                            }
                        }
                        // Our end point is actually one index above the current index
                        mEndPoint = k + 1;
                    }
                }
                // Add the last set of columns
                mColumnPoints.Add(new int[] { mStartPoint, mEndPoint });
                mColumnPointsWidth.Add(mTempWidth);
                mColumnPoint = 0;
            }
        }

        // The funtion that print the title, page number, and the header row
        private static void DrawHeader(Graphics g)
        {
            CurrentY = (float)TopMargin;

            // Printing the page number (if isWithPaging is set to true)
            if (IsWithPaging)
            {
                PageNumber++;
                string PageString = "Page " + PageNumber.ToString();

                StringFormat PageStringFormat = new StringFormat();
                PageStringFormat.Trimming = StringTrimming.Word;
                PageStringFormat.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
                PageStringFormat.Alignment = StringAlignment.Far;

                Font PageStringFont = new Font("Tahoma", 8, FontStyle.Regular, GraphicsUnit.Point);

                RectangleF PageStringRectangle = new RectangleF((float)LeftMargin, CurrentY, (float)PageWidth - (float)RightMargin - (float)LeftMargin, g.MeasureString(PageString, PageStringFont).Height);

                g.DrawString(PageString, PageStringFont, new SolidBrush(Color.Black), PageStringRectangle, PageStringFormat);

                CurrentY += g.MeasureString(PageString, PageStringFont).Height;
            }

            // Printing the title (if IsWithTitle is set to true)
            if (IsWithTitle)
            {
                StringFormat TitleFormat = new StringFormat();
                TitleFormat.Trimming = StringTrimming.Word;
                TitleFormat.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
                if (IsCenterOnPage)
                    TitleFormat.Alignment = StringAlignment.Center;
                else
                    TitleFormat.Alignment = StringAlignment.Near;

                RectangleF TitleRectangle = new RectangleF((float)LeftMargin, CurrentY, (float)PageWidth - (float)RightMargin - (float)LeftMargin, g.MeasureString(TheTitleText, TheTitleFont).Height);

                g.DrawString(TheTitleText, TheTitleFont, new SolidBrush(TheTitleColor), TitleRectangle, TitleFormat);

                CurrentY += g.MeasureString(TheTitleText, TheTitleFont).Height;
            }

            // Calculating the starting x coordinate that the printing process will start from
            float CurrentX = (float)LeftMargin;
            if (IsCenterOnPage)
                CurrentX += (((float)PageWidth - (float)RightMargin - (float)LeftMargin) - mColumnPointsWidth[mColumnPoint]) / 2.0F;

            // Setting the HeaderFore style
            Color HeaderForeColor = TheDataGridView.ColumnHeadersDefaultCellStyle.ForeColor;
            if (HeaderForeColor.IsEmpty) // If there is no special HeaderFore style, then use the default DataGridView style
                HeaderForeColor = TheDataGridView.DefaultCellStyle.ForeColor;
            SolidBrush HeaderForeBrush = new SolidBrush(HeaderForeColor);

            // Setting the HeaderBack style
            Color HeaderBackColor = TheDataGridView.ColumnHeadersDefaultCellStyle.BackColor;
            if (HeaderBackColor.IsEmpty) // If there is no special HeaderBack style, then use the default DataGridView style
                HeaderBackColor = TheDataGridView.DefaultCellStyle.BackColor;
            SolidBrush HeaderBackBrush = new SolidBrush(HeaderBackColor);

            // Setting the LinePen that will be used to draw lines and rectangles (derived from the GridColor property of the DataGridView control)
            Pen TheLinePen = new Pen(TheDataGridView.GridColor, 1);

            // Setting the HeaderFont style
            Font HeaderFont = TheDataGridView.ColumnHeadersDefaultCellStyle.Font;
            if (HeaderFont == null) // If there is no special HeaderFont style, then use the default DataGridView font style
                HeaderFont = TheDataGridView.DefaultCellStyle.Font;

            // Calculating and drawing the HeaderBounds        
            RectangleF HeaderBounds = new RectangleF(CurrentX, CurrentY, mColumnPointsWidth[mColumnPoint], RowHeaderHeight);
            g.FillRectangle(HeaderBackBrush, HeaderBounds);

            // Setting the format that will be used to print each cell of the header row
            StringFormat CellFormat = new StringFormat();
            CellFormat.Trimming = StringTrimming.Word;
            CellFormat.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.LineLimit | StringFormatFlags.NoClip;

            // Printing each visible cell of the header row
            RectangleF CellBounds;
            float ColumnWidth;
            for (int i = (int)mColumnPoints[mColumnPoint].GetValue(0); i < (int)mColumnPoints[mColumnPoint].GetValue(1); i++)
            {
                if (!TheDataGridView.Columns[i].Visible) continue; // If the column is not visible then ignore this iteration

                ColumnWidth = ColumnsWidth[i];

                // Check the CurrentCell alignment and apply it to the CellFormat
                if (TheDataGridView.ColumnHeadersDefaultCellStyle.Alignment.ToString().Contains("Right"))
                    CellFormat.Alignment = StringAlignment.Far;
                else if (TheDataGridView.ColumnHeadersDefaultCellStyle.Alignment.ToString().Contains("Center"))
                    CellFormat.Alignment = StringAlignment.Center;
                else
                    CellFormat.Alignment = StringAlignment.Near;

                CellBounds = new RectangleF(CurrentX, CurrentY, ColumnWidth, RowHeaderHeight);

                // Printing the cell text
                g.DrawString(TheDataGridView.Columns[i].HeaderText, HeaderFont, HeaderForeBrush, CellBounds, CellFormat);

                // Drawing the cell bounds
                if (TheDataGridView.RowHeadersBorderStyle != DataGridViewHeaderBorderStyle.None) // Draw the cell border only if the HeaderBorderStyle is not None
                    g.DrawRectangle(TheLinePen, CurrentX, CurrentY, ColumnWidth, RowHeaderHeight);

                CurrentX += ColumnWidth;
            }

            CurrentY += RowHeaderHeight;
        }

        // The function that print a bunch of rows that fit in one page
        // When it returns true, meaning that there are more rows still not printed, so another PagePrint action is required
        // When it returns false, meaning that all rows are printed (the CureentRow parameter reaches the last row of the DataGridView control) and no further PagePrint action is required
        private static bool DrawRows(Graphics g)
        {
            // Setting the LinePen that will be used to draw lines and rectangles (derived from the GridColor property of the DataGridView control)
            Pen TheLinePen = new Pen(TheDataGridView.GridColor, 1);

            // The style paramters that will be used to print each cell
            Font RowFont;
            Color RowForeColor;
            Color RowBackColor;
            SolidBrush RowForeBrush;
            SolidBrush RowBackBrush;
            SolidBrush RowAlternatingBackBrush;

            // Setting the format that will be used to print each cell
            StringFormat CellFormat = new StringFormat();
            CellFormat.Trimming = StringTrimming.Word;
            CellFormat.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.LineLimit;

            // Printing each visible cell
            RectangleF RowBounds;
            float CurrentX;
            float ColumnWidth;
            while (CurrentRow < TheDataGridView.Rows.Count)
            {
                if (TheDataGridView.Rows[CurrentRow].Visible) // Print the cells of the CurrentRow only if that row is visible
                {
                    // Setting the row font style
                    RowFont = TheDataGridView.Rows[CurrentRow].DefaultCellStyle.Font;
                    if (RowFont == null) // If the there is no special font style of the CurrentRow, then use the default one associated with the DataGridView control
                        RowFont = TheDataGridView.DefaultCellStyle.Font;

                    // Setting the RowFore style
                    RowForeColor = TheDataGridView.Rows[CurrentRow].DefaultCellStyle.ForeColor;
                    if (RowForeColor.IsEmpty) // If the there is no special RowFore style of the CurrentRow, then use the default one associated with the DataGridView control
                        RowForeColor = TheDataGridView.DefaultCellStyle.ForeColor;
                    RowForeBrush = new SolidBrush(RowForeColor);

                    // Setting the RowBack (for even rows) and the RowAlternatingBack (for odd rows) styles
                    RowBackColor = TheDataGridView.Rows[CurrentRow].DefaultCellStyle.BackColor;
                    if (RowBackColor.IsEmpty) // If the there is no special RowBack style of the CurrentRow, then use the default one associated with the DataGridView control
                    {
                        RowBackBrush = new SolidBrush(TheDataGridView.DefaultCellStyle.BackColor);
                        RowAlternatingBackBrush = new SolidBrush(TheDataGridView.AlternatingRowsDefaultCellStyle.BackColor);
                    }
                    else // If the there is a special RowBack style of the CurrentRow, then use it for both the RowBack and the RowAlternatingBack styles
                    {
                        RowBackBrush = new SolidBrush(RowBackColor);
                        RowAlternatingBackBrush = new SolidBrush(RowBackColor);
                    }

                    // Calculating the starting x coordinate that the printing process will start from
                    CurrentX = (float)LeftMargin;
                    if (IsCenterOnPage)
                        CurrentX += (((float)PageWidth - (float)RightMargin - (float)LeftMargin) - mColumnPointsWidth[mColumnPoint]) / 2.0F;

                    // Calculating the entire CurrentRow bounds                
                    RowBounds = new RectangleF(CurrentX, CurrentY, mColumnPointsWidth[mColumnPoint], RowsHeight[CurrentRow]);

                    // Filling the back of the CurrentRow
                    if (CurrentRow % 2 == 0)
                        g.FillRectangle(RowBackBrush, RowBounds);
                    else
                        g.FillRectangle(RowAlternatingBackBrush, RowBounds);

                    // Printing each visible cell of the CurrentRow                
                    for (int CurrentCell = (int)mColumnPoints[mColumnPoint].GetValue(0); CurrentCell < (int)mColumnPoints[mColumnPoint].GetValue(1); CurrentCell++)
                    {
                        if (!TheDataGridView.Columns[CurrentCell].Visible) continue; // If the cell is belong to invisible column, then ignore this iteration

                        // Check the CurrentCell alignment and apply it to the CellFormat
                        if (TheDataGridView.Columns[CurrentCell].DefaultCellStyle.Alignment.ToString().Contains("Right"))
                            CellFormat.Alignment = StringAlignment.Far;
                        else if (TheDataGridView.Columns[CurrentCell].DefaultCellStyle.Alignment.ToString().Contains("Center"))
                            CellFormat.Alignment = StringAlignment.Center;
                        else
                            CellFormat.Alignment = StringAlignment.Near;

                        ColumnWidth = ColumnsWidth[CurrentCell];
                        RectangleF CellBounds = new RectangleF(CurrentX, CurrentY, ColumnWidth, RowsHeight[CurrentRow]);

                        // Printing the cell text
                        g.DrawString(TheDataGridView.Rows[CurrentRow].Cells[CurrentCell].EditedFormattedValue.ToString(), RowFont, RowForeBrush, CellBounds, CellFormat);

                        // Drawing the cell bounds
                        if (TheDataGridView.CellBorderStyle != DataGridViewCellBorderStyle.None) // Draw the cell border only if the CellBorderStyle is not None
                            g.DrawRectangle(TheLinePen, CurrentX, CurrentY, ColumnWidth, RowsHeight[CurrentRow]);

                        CurrentX += ColumnWidth;
                    }
                    CurrentY += RowsHeight[CurrentRow];

                    // Checking if the CurrentY is exceeds the page boundries
                    // If so then exit the function and returning true meaning another PagePrint action is required
                    if ((int)CurrentY > (PageHeight - TopMargin - BottomMargin))
                    {
                        CurrentRow++;
                        return true;
                    }
                }
                CurrentRow++;
            }

            CurrentRow = 0;
            mColumnPoint++; // Continue to print the next group of columns

            if (mColumnPoint == mColumnPoints.Count) // Which means all columns are printed
            {
                mColumnPoint = 0;
                return false;
            }
            else
                return true;
        }

        // The method that calls all other functions
      public static bool DrawDataGridView(Graphics g)
        {
            try
            {
                Calculate(g);
                DrawHeader(g);
                bool bContinue = DrawRows(g);
                return bContinue;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Operation failed: " + ex.Message.ToString(), Application.ProductName + " - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }



}
