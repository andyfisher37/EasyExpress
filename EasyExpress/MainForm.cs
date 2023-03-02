/// EasyExpress - making report tool for police department on evereymonth form "Kadry-Express"
/// using MS SQL Server database for storing data

using DevExpress.XtraRichEdit.API.Native;
using Janus.Windows.GridEX;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace EasyExpress
{
	public partial class MainForm : Form
    {
        public static int[] pos = new int[42];  // Массив расчетных позиций
        public static int[,] express = new int[1000, 42]; // Массив всего отчета
        // Список используемых ДГ до ноября 2016
        //public List<int> DG = new List<int> { 10, 13, 20, 30, 40, 50, 65, 87, 92, 117, 170, 191, 200, 210, 287, 289, 300, 301, 350, 360, 362, 399, 400, 430, 829, 879, 900, 901, 971, 972, 978, 979 };
        // Список используемых ДГ актуальных
        public List<int> DG = new List<int> { 10, 13, 20, 30, 40, 50, 65, 87, 92, 117, 170, 191, 200, 210, 287, 289, 300, 301, 350, 360, 362, 400, 430, 900, 901, 971, 972, 978, 979 };
        // Перечисление ДГ в порядке записываемом в отчет
        public List<string> repDG = new List<string> { "010", "020", "030", "971", "972", "978", "979", "092", "117", "170", "210", "200", "191", "087", "289", "287", "350", "362", "040", "050", "065", "013", "360", "400", "430" };
        public int count_repDG = 25; // Количество ДГ в отчете
        public static string current_dg_code;   // Текущая должностная группа
        public static string current_sql_text;  // Текущий текст запроса ДГ
        public static string current_dg_name;   // Наименование текущей ДГ
        public static DateTime calc_date = Convert.ToDateTime(DateTime.Now.ToShortDateString()); // Дата расчета
        public static string calc_year = calc_date.Year.ToString();            // Год расчета
        public static string calc_month = DigMonth(calc_date);          // Месяц расчета
        public static DateTime res_date;
        public static string kadryConnection;
        public static string sqlConnection;


        public MainForm()
        {
            InitializeComponent();
        }

        // Преобразование цифры МЕСЯЦА в двухзначный вид
        public static string DigMonth(DateTime d)
        {
            if (d.Month < 10) return "0" + d.Month.ToString();
            else return d.Month.ToString();
        }

        // Инициализация расчетного массива
        public void InitPosArray(int maxinit)
        {
            for (int i = 0; i < maxinit; i++) pos[i] = 0;
        }

        // Инициализация всего отчета 
        public void InitAllExpress(int[,] dim)
        {
            for (int i = 0; i < 1000; i++)
            {
                for (int j = 0; j < 42; j++) dim[i, j] = 0;
            }
        }

        // Сохранение PosArray в Express
        public void SavePos2Express(string code)
        {
            int n = Convert.ToInt16(code);

            for (int i = 0; i < 42; i++) express[n, i] = pos[i]; 
        }

        // Сохранение всего отчета в БД
        public void SaveAllExpressToBase()
        {
            for (int i = 0; i < 1000; i++)
            {
                if (DG.Contains(i))
                {
                    if ( i < 100 ) SaveDGtoBase("0" + i.ToString(), calc_date.ToShortDateString());
                    else SaveDGtoBase(i.ToString(), calc_date.ToShortDateString());
                }
            }
            MessageBox.Show("Отчет сохранен!");
        }

        // Загрузка данных отчета из БД по дате
        public void LoadAllExpressFromBase(DateTime date, int[,] dim)
        {
            InitAllExpress(dim);

            DataTable dt = DataProvider._getDataSQL(sqlConnection,String.Format("SELECT * FROM Express_2012 WHERE calc_date = '{0}' ORDER BY code_dg", DateToStr(date)));
            DataRowCollection rc = dt.Rows;
            dt.Dispose();

            if (rc.Count > 0)
            {

                for (int i = 0; i < rc.Count; i++)
                {
                    int code = Convert.ToInt16(rc[i]["code_dg"]);
                    for (int j = 1; j < 42; j++) dim[code, j] = Convert.ToInt16(rc[i]["pos_" + j.ToString()]);
                }

                rc.Clear();
                MessageBox.Show("Отчет загружен!");

                calc_on_date.Value = date;
                script_date.Value = date;
            }
            else MessageBox.Show("Ошибка загрузки данных!");
        }

        // Проверка между должностными группами
        public bool CheckInterDG()
        {
            int cnt = 0;
            int [] sum1 = new int[42];
            int [] sum2 = new int[42];

            debug.Items.Clear();

            for (int i = 1; i < 42; i++)
            {
                // 020+030=971+972+978+979+360+362
                sum1[i] = express[20,i] + express[30,i];
                sum2[i] = express[971, i] + express[972, i] + express[978, i] + express[979, i] + express[360, i] + express[362, i]; ;
                if (sum1[i] != sum2[i])
                {
                    cnt++;
                    debug.Items.Add(String.Format("Не бъется сумма: [020]+[030] = [971]+[972]+[978]+[979]+[360]+[362] в позиции {0} | {1} = {2}", i,sum1[i],sum2[i]));
                }
                
                // 020>=971+978+360
                sum1[i] = express[971,i] + express[978,i] + express[360,i];
                if ( express[20,i] < sum1[i] )
                {
                    cnt++;
                    debug.Items.Add(String.Format("Не выполнено условие: [020] >= [971]+[978]+[360] в позиции {0} | {1} < {2} ", i, express[20,i], sum1[i]));
                }
                
                // 030>=972+979+362
                sum1[i] = express[972,i] + express[979,i] + express[362,i];
                if ( express[30,i] < sum1[i] )
                {
                    cnt++;
                    debug.Items.Add(String.Format("Не выполнено условие: [030] >= [972]+[979]+[362] в позиции {0} | {1} < {2}", i, express[30,i], sum1[i]));
                }
                
                //170+191+210+350<=971
                sum1[i] = express[170,i] + express[191,i] + express[210,i] + express[350,i] + express[350,i];
                if ( sum1[i] > express[971,i] )
                {
                    cnt++;
                    debug.Items.Add(String.Format("Не выполнено условие: [971] >= [170]+[191]+[210]+[350] в позиции {0} | {1} > {2} ", i, sum1[i], express[971,i]));
                }
                
                // 289>=287
                if ( express[289,i] < express[287,i] ) { cnt++; debug.Items.Add(String.Format("Не выполнено условие: [289] >= [287] в позиции {0} | {1} < {2}", i, express[289,i], express[287,i])); }

                // 200+087+289+399+829+879<=971+972
                sum1[i] = express[200,i] + express[87,i] + express[289,i] + express[399,i] + express[829,i] + express[879,i];
                sum2[i] = express[971,i] + express[972,i];
                if (sum1[i] > sum2[i])
                {
                    cnt++;
                    debug.Items.Add(String.Format("Не выполняется условие: [200]+[087]+[289]+[399]+[829]+[879]<=[971]+[972] в позиции {0} | {1} > {2}", i, sum1[i], sum2[i]));
                }
            }
            
            if (cnt > 0) return false;
            else return true;
        }

        // Проверка должностной группы
        public bool CheckDG(string code)
        {
            int cnt = 0; // счетчик ошибок

            debug.Items.Clear();

            if (pos[1] <= 0)
            {
                debug.Items.Add("Позиция [01] не может быть меньше или равна нулю\n");
                cnt++;
            }
            if (pos[2] > pos[1])
            {
                debug.Items.Add(String.Format("Замещение [02] не может быть больше штатной численности [01] | {0} > {1}\n",pos[2],pos[1]));
                cnt++;
            }
            if ((pos[4] + pos[5] + pos[6]) != pos[3])
            {
                debug.Items.Add(String.Format("Не бъется сумма принятых [03] = [04]+[05]+[06] | {0} != {1} + {2} + {3}\n",pos[3],pos[4],pos[5],pos[6]));
                cnt++;
            }
            if (pos[8] > pos[7]) 
            {
                debug.Items.Add(String.Format("Всего уволено должно быть больше или равно количеству уволенных до 1 года | {0} > {1}\n", pos[8], pos[7]));
                cnt++;
            }
            if (pos[11] > pos[7]) 
            {
                debug.Items.Add(String.Format("Уволенных из резерва [11] должно быть меньше или равно всего уволено [07] | {0} > {1}\n", pos[11], pos[7]));
                cnt++;
            }
            if (pos[15] > pos[14])
            {
                debug.Items.Add(String.Format("Сменяемость по ротации должна быть меньше или равна общей сменяемости | {0] > {1}\n", pos[15] > pos[14]));
                cnt++;
            }
            // -----------------------------------------------------------------------------------------------------
            if (pos[17] > pos[16])
            {
                debug.Items.Add(String.Format("Погибших всего должно быть больше или равно погибшим при исп.обязанностей | {0} > {1}\n", pos[17], pos[16]));
                cnt++;
            }
            if ((pos[18] + pos[19] + pos[20] + pos[21] + pos[22] + pos[23] + pos[28]) != pos[17])
            {
                debug.Items.Add("Не бъется сумма погибших при исполнении...\n");
                cnt++;
            }
            if (pos[23] < (pos[24] + pos[25] + pos[26] + pos[27]))
            {
                debug.Items.Add("Не бъется сумма погибших 'в результате непроффес.действий'\n");
                cnt++;
            }
            if (pos[30] > pos[29]) 
            {
                debug.Items.Add("Ранено при исполнении должно быть меньше или равно всего раненным\n");
                cnt++;
            }
            if ((pos[31] + pos[32] + pos[33] + pos[34] + pos[35] + pos[36] + pos[41]) != pos[30])
            {
                debug.Items.Add("Не бъется сумма ранненных при исполнении...\n"); 
                cnt++;
            }
            if ((pos[37] + pos[38] + pos[39] + pos[40]) < pos[36]) 
            {
                debug.Items.Add("Не бъется сумма раненных 'в результате непроффес.действий'\n");
                cnt++;
            }
 
            if (cnt > 0) return false;
            else return true;
        }

        // Расчет должностной группы ( код, текс запроса )
        public void CalcDG(string code, string sql_text)
        {
            if (checkPos.Checked) InitPosArray(16);
            else InitPosArray(42);

            progress.Maximum = 15;
            progress.Value = 0;

            // Должностей по штату
            pos[1] = DataProvider._getDataSQLs(kadryConnection,String.Format("SELECT SUM(STAVKA_DLZ) FROM AAQQ {0} AND DATA_SOKR IS NULL", sql_text));
            progress.Value++;

            // из них замещено
            if (code != "010" && code != "013")
            {
                // Если аттестованные считаем по фамилиям
                pos[2] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM AAQQ {0} AND FAMILIYA <> ''", sql_text));
                progress.Value++;
            }
            else
            {
                // Если в/н считаем по ставкам
                pos[2] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT SUM(STAVKA_PRS) FROM AAQQ {0} AND FAMILIYA <> ''", sql_text));
                progress.Value++;
            }
            
            // Принято за отчетный период 
            //(нарастающим итогом с начала года)
            pos[3] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM PRIEM {0} AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            progress.Value++;
            
            // В том числе вновь принятых (в т.ч. стажеры и выпускники ОУ МВД)
            //(нарастающим итогом с начала года)
            pos[4] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM PRIEM {0} AND (KAT_POST NOT IN (101,102,104) OR KAT_POST IS NULL) AND (DAT_REG >= {1} AND DAT_REG <= {2})", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            progress.Value++;

            // В том числе прибывших из органов внутренних дел (др.регионы - только из системы МВД !)
            //(нарастающим итогом с начала года)
            pos[5] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM PRIEM {0} AND KAT_POST=101 AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            progress.Value++;

            // В том числе прибывших из подразделений Центрального аппарата МВД
            //(нарастающим итогом с начала года)
            pos[6] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM PRIEM {0} AND KAT_POST=104 AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            progress.Value++;

            // Уволено
            //(нарастающим итогом с начала года)
            pos[7] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM ARCHIVE {0} AND PRICH_UV <> '1021' AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            progress.Value++;

            // Из них на 1 году службы
            //(нарастающим итогом с начала года)
            pos[8] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM ARCHIVE {0} AND PRICH_UV <> '1021' AND ZVANIE <> 99 AND ( ( (SL_RANE_OT IS NULL)  AND  ( (DATA_UVOL - DATA_POST)  < 365) )  OR  ( (SL_RANE_OT IS NOT NULL)  AND  ( ( (DATA_UVOL - DATA_POST)  +  (SL_RANE_DO - SL_RANE_OT) )  < 365) ) )  AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            progress.Value++;

            // Откомандировано
            //(нарастающим итогом с начала года)
            // Код выбытия - 104 - только для "хитрых"
            pos[9] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM VYEZD {0} AND KOD_VYBYL NOT IN (104) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            progress.Value++;

            // Находятся в распоряжениии ОВД за отчетный месяц
            pos[10] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM RESERV {0} AND PRIMECHAN NOT LIKE 'декрет%' AND DATA_ZACH >= {1} AND DATA_ZACH <= {2}", sql_text, res_date.ToOADate(), calc_date.ToOADate()));
            progress.Value++;

            // Добавить фильтр по датам (ЗА МЕСЯЦ!!!)
            // Движение по распоряженцам...
            if (code != "010" && code != "013" && code != "971" && code != "972" && code != "978" && code != "979")
            {
                DataTable dt = DataProvider._getDataSQL(kadryConnection, String.Format("SELECT DISTINCT KEY_POSL FROM POSL_SPI {0} AND STATUS IN ('6000') AND DATA_OT >= {1}", sql_text, Convert.ToDateTime("01." + calc_month + "." + calc_year).ToOADate()));
                DataRowCollection rc = dt.Rows;
                dt.Dispose();
                if (rc.Count > 0)
                {
                    string keys = "";
                    for (int n = 0; n < rc.Count; n++)
                    {
                        if (n == 0) keys += rc[n]["KEY_POSL"].ToString();
                        else keys += "," + rc[n]["KEY_POSL"].ToString();
                    }
                    rc.Clear();
                    // Уволено сотрудников, находящихся в распоряжении ОВД за отчетный месяц
                    pos[11] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM ARCHIVE WHERE KEY_1 IN ({0})", keys));
                    progress.Value++;

                    // Назначено сотрудников на должность в данном подразделении, из числа находящихся в распоряжении ОВД за отчетный месяц
                    pos[12] = 0;
                    progress.Value++;

                    // Откомандировано в другое подразделение, из числа находящихся в распоряжении ОВД, за отчетный месяц
                    pos[13] = pos[11] = DataProvider._getDataSQLs(kadryConnection, String.Format("SELECT COUNT(FAMILIYA) FROM VYEZD WHERE KEY_1 IN ({0})", keys)); ;
                    progress.Value++;
                }
            }
            else
            {
                // Уволено сотрудников, находящихся в распоряжении ОВД за отчетный месяц
                pos[11] = 0;
                progress.Value++;

                // Назначено сотрудников на должность в данном подразделении, из числа находящихся в распоряжении ОВД за отчетный месяц
                pos[12] = 0;
                progress.Value++;

                // Откомандировано в другое подразделение, из числа находящихся в распоряжении ОВД, за отчетный месяц
                pos[13] = 0;
                progress.Value++;
            }

            // Сменилось руководителей (только ДГ [092] и [117])
            // начальники территориальных на районном уровне...
            //(нарастающим итогом с начала года)
            // ! без оргштатных
            if (current_dg_code == "092")
            {
                // Сменилось всего (все кто ушел с должностей нач.ТОВД)
                // Выбираем ключи всех сотрудников (действующих) у кого в текущем году сменилась должность
                DataTable dt = DataProvider._getDataSQL(kadryConnection, String.Format("select key_posl from posl_spi where data_ot BETWEEN {0} AND {1} and key_posl in (select distinct key_1 from aaqq where key_1 <> 0 and dolznost < '200000') order by key_posl", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                // получили коллекцию ключей 
                DataRowCollection rc = dt.Rows;
                dt.Dispose();
                if (rc.Count > 0)
                {
                    // По каждому сотруднику смотрим 2 "верхние" записи в послужном (отсортированном по датам в порядке убывания)...
                    for (int i = 0; i < rc.Count; i++)
                    {
                        string key = rc[i]["key_posl"].ToString();
                        DataTable tmp = DataProvider._getDataSQL(kadryConnection, String.Format("select top(2) sluzba from posl_spi where key_posl = {0} order by data_ot desc", key));
                        DataRowCollection sl = tmp.Rows;
                        if (tmp.Rows.Count == 2)
                        {
                            // если в предыдущей (из 2) записи служба 18 (и текущая другая) - значит была сменяемость.
                            if (sl[1]["sluzba"].ToString() == "16" && sl[0]["sluzba"].ToString() != "16") pos[14]++;
                        }
                        tmp.Clear();
                        Application.DoEvents();
                    }

                    rc.Clear();

                    // добавляем уволенных, зачисленных в распоряжение и откомандированных
                    pos[14] += DataProvider._getDataSQLs(kadryConnection,
                               String.Format("select a = (select count(key_1) from archive where sluzba in (16) and DAT_REG BETWEEN {0} AND {1}) + " +
                                             "(select count(key_1) from RESERV where sluzba in (16) and DATA_ZACH BETWEEN {0} AND {1}) + " +
                                             "(select count(key_1) from VYEZD where sluzba in (16) and DATA_UVOL BETWEEN {0} AND {1})", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    progress.Value++;

                    // Сменилось руководителей в порядке ротации (по решению През.РФ)
                    pos[15] = 0;
                    progress.Value++;
                }
                else
                {
                    pos[14] = 0;
                    progress.Value++;
                    pos[15] = 0;
                    progress.Value++;
                }
            }
            // заместители начальников территориальных на районном уровне
            else if (current_dg_code == "117")
            {
                // Сменилось всего (все кто ушел с должностей зам.нач.ТОВД)
                // Выбираем ключи всех сотрудников (действующих) у кого в текущем году сменилась должность
                DataTable dt = DataProvider._getDataSQL(kadryConnection, String.Format("select key_posl from posl_spi where data_ot BETWEEN {0} AND {1} and key_posl in (select distinct key_1 from aaqq where key_1 <> 0 and dolznost < '500000') order by key_posl", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                // получили коллекцию ключей 
                DataRowCollection rc = dt.Rows;
                dt.Dispose();
                if (rc.Count > 0)
                {
                    // По каждому сотруднику смотрим 2 "верхние" записи в послужном (отсортированном по датам в порядке убывания)...
                    for (int i = 0; i < rc.Count; i++)
                    {
                        string key = rc[i]["key_posl"].ToString();
                        DataTable tmp = DataProvider._getDataSQL(kadryConnection, String.Format("select top(2) sluzba, dolznost from posl_spi where key_posl = {0} order by data_ot desc", key));
                        DataRowCollection sm = tmp.Rows;
                        if (tmp.Rows.Count == 2)
                        {
                            // если в предыдущей (из 2) записи служба 17,18 (и текущая другая) - значит была сменяемость.
                            // (sluzba in (17, 18) or DOLZNOST IN('100493', '100587', '100586'))
                            // следствие по должностям '100587','100586','100667'
                            if (sm[1]["sluzba"].ToString() == "17" && sm[0]["sluzba"].ToString() != "17") pos[14]++;
                            if (sm[1]["sluzba"].ToString() == "18" && sm[0]["sluzba"].ToString() != "18") pos[14]++;
                            if (sm[1]["dolznost"].ToString() == "100587" && sm[0]["dolznost"].ToString() != "100587") pos[14]++;
                            if (sm[1]["dolznost"].ToString() == "100586" && sm[0]["dolznost"].ToString() != "100586") pos[14]++;
                            if (sm[1]["dolznost"].ToString() == "100667" && sm[0]["dolznost"].ToString() != "100667") pos[14]++;
                        }
                        tmp.Clear();
                        Application.DoEvents();
                    }

                    rc.Clear();

                    // добавляем уволенных, зачисленных в распоряжение и откомандированных
                    pos[14] += DataProvider._getDataSQLs(kadryConnection,
                               String.Format("select a = (select count(key_1) from archive where (sluzba in (17,18) or DOLZNOST IN ('100587','100586','100667')) and DAT_REG BETWEEN {0} AND {1}) + " +
                                             "(select count(key_1) from RESERV where (sluzba in (17,18) or DOLZNOST IN ('100587','100586','100667')) and DATA_ZACH BETWEEN {0} AND {1}) + " +
                                             "(select count(key_1) from VYEZD where (sluzba in (17,18) or DOLZNOST IN ('100587','100586','100667')) and DATA_UVOL BETWEEN {0} AND {1})", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    progress.Value++;

                    // Сменилось руководителей в порядке ротации (по решению През.РФ)
                    pos[15] = 0;
                    progress.Value++;
                }
                else
                {
                    pos[14] = 0;
                    progress.Value++;
                    pos[15] = 0;
                    progress.Value++;
                }

            }

                //MessageBox.Show("Well done!");

                progress.Value = 0;

            SavePos2Express(code);
        }

        // Расчет некомплекта по ДГ
        private double CalcNek(int code)
        {
            double s = Convert.ToDouble(express[code, 1]);
            double z = Convert.ToDouble(express[code, 2]);
            return Math.Round((s - z) / s * 100, 1);
        }

        // Заполнение формы редактирования
        public void FillEditPanel(string code, int swich)
        {
            if ( swich == 1 && mainTabControl.SelectedTabPageIndex != 1) mainTabControl.SelectedTabPageIndex = 1;

            int icode = Convert.ToInt16(code);

            Edit_groupBox.Text = String.Format("[{0}] - {1}",current_dg_code, current_dg_name);
            nek_info.Text = String.Format("Некомплект: {0} ед. ({1}) ", express[icode,1]-express[icode,2], CalcNek(icode));

            edit_pos1.Text = express[icode,1].ToString();
            edit_pos2.Text = express[icode,2].ToString();
            edit_pos3.Text = express[icode,3].ToString();
            edit_pos4.Text = express[icode,4].ToString();
            edit_pos5.Text = express[icode,5].ToString();
            edit_pos6.Text = express[icode,6].ToString();
            edit_pos7.Text = express[icode,7].ToString();
            edit_pos8.Text = express[icode,8].ToString();
            edit_pos9.Text = express[icode,9].ToString();
            edit_pos10.Text = express[icode,10].ToString();
            edit_pos11.Text = express[icode,11].ToString();
            edit_pos12.Text = express[icode,12].ToString();
            edit_pos13.Text = express[icode,13].ToString();
            edit_pos14.Text = express[icode,14].ToString();
            edit_pos15.Text = express[icode,15].ToString();

            edit_pos16.Text = express[icode, 16].ToString();
            edit_pos17.Text = express[icode, 17].ToString();
            edit_pos18.Text = express[icode, 18].ToString();
            edit_pos19.Text = express[icode, 19].ToString();
            edit_pos20.Text = express[icode, 20].ToString();
            edit_pos21.Text = express[icode, 21].ToString();
            edit_pos22.Text = express[icode, 22].ToString();
            edit_pos23.Text = express[icode, 23].ToString();
            edit_pos24.Text = express[icode, 24].ToString();
            edit_pos25.Text = express[icode, 25].ToString();
            edit_pos26.Text = express[icode, 26].ToString();
            edit_pos27.Text = express[icode, 27].ToString();
            edit_pos28.Text = express[icode, 28].ToString();
            edit_pos29.Text = express[icode, 29].ToString();
            edit_pos30.Text = express[icode, 30].ToString();
            edit_pos31.Text = express[icode, 31].ToString();
            edit_pos32.Text = express[icode, 32].ToString();
            edit_pos33.Text = express[icode, 33].ToString();
            edit_pos34.Text = express[icode, 34].ToString();
            edit_pos35.Text = express[icode, 35].ToString();
            edit_pos36.Text = express[icode, 36].ToString();
            edit_pos37.Text = express[icode, 37].ToString();
            edit_pos38.Text = express[icode, 38].ToString();
            edit_pos39.Text = express[icode, 39].ToString();
            edit_pos40.Text = express[icode, 40].ToString();
            edit_pos41.Text = express[icode, 41].ToString();
        }

        // Обновление данных в Pos из панели редактирования
        public void RefreshPosFromEditPanel()
        {
            pos[1] = Convert.ToInt16(edit_pos1.Text);
            pos[2] = Convert.ToInt16(edit_pos2.Text);
            pos[3] = Convert.ToInt16(edit_pos3.Text);
            pos[4] = Convert.ToInt16(edit_pos4.Text);
            pos[5] = Convert.ToInt16(edit_pos5.Text);
            pos[6] = Convert.ToInt16(edit_pos6.Text);
            pos[7] = Convert.ToInt16(edit_pos7.Text);
            pos[8] = Convert.ToInt16(edit_pos8.Text);
            pos[9] = Convert.ToInt16(edit_pos9.Text);
            pos[10] = Convert.ToInt16(edit_pos10.Text);
            pos[11] = Convert.ToInt16(edit_pos11.Text);
            pos[12] = Convert.ToInt16(edit_pos12.Text);
            pos[13] = Convert.ToInt16(edit_pos13.Text);
            pos[14] = Convert.ToInt16(edit_pos14.Text);
            pos[15] = Convert.ToInt16(edit_pos15.Text);

            pos[16] = Convert.ToInt16(edit_pos16.Text);
            pos[17] = Convert.ToInt16(edit_pos17.Text);
            pos[18] = Convert.ToInt16(edit_pos18.Text);
            pos[19] = Convert.ToInt16(edit_pos19.Text);
            pos[20] = Convert.ToInt16(edit_pos20.Text);
            pos[21] = Convert.ToInt16(edit_pos21.Text);
            pos[22] = Convert.ToInt16(edit_pos22.Text);
            pos[23] = Convert.ToInt16(edit_pos23.Text);
            pos[24] = Convert.ToInt16(edit_pos24.Text);
            pos[25] = Convert.ToInt16(edit_pos25.Text);
            pos[26] = Convert.ToInt16(edit_pos26.Text);
            pos[27] = Convert.ToInt16(edit_pos27.Text);
            pos[28] = Convert.ToInt16(edit_pos28.Text);
            pos[29] = Convert.ToInt16(edit_pos29.Text);
            pos[30] = Convert.ToInt16(edit_pos30.Text);
            pos[31] = Convert.ToInt16(edit_pos31.Text);
            pos[32] = Convert.ToInt16(edit_pos32.Text);
            pos[33] = Convert.ToInt16(edit_pos33.Text);
            pos[34] = Convert.ToInt16(edit_pos34.Text);
            pos[35] = Convert.ToInt16(edit_pos35.Text);
            pos[36] = Convert.ToInt16(edit_pos36.Text);
            pos[37] = Convert.ToInt16(edit_pos37.Text);
            pos[38] = Convert.ToInt16(edit_pos38.Text);
            pos[39] = Convert.ToInt16(edit_pos39.Text);
            pos[40] = Convert.ToInt16(edit_pos40.Text);
            pos[41] = Convert.ToInt16(edit_pos41.Text);
        }

        // Печать в отчет 
        public void Write2Report(string pos, string val)
        {
            DocumentPosition docpos = report.Document.Bookmarks[pos].Range.Start;
            report.Document.CaretPosition = docpos;
            report.ScrollToCaret();
            report.Document.InsertText(docpos, val);
        }

        // Преобразование числа месяца в строку
        private string Month2String(int m)
        {
            switch(m)
            {
                case 1: return "январь";
                case 2: return "февраль";
                case 3: return "март";
                case 4: return "апрель";
                case 5: return "май";
                case 6: return "июнь";
                case 7: return "июль";
                case 8: return "август";
                case 9: return "сентябрь";
                case 10: return "октябрь";
                case 11: return "ноябрь";
                case 12: return "декабрь";
                default: return "";
            }
        }

        // Преобразование даты в строку типа: yyyy-mm-dd
        private string DateToStr(DateTime date)
        {
            string res = date.Year.ToString() + "-" + date.Month.ToString() + "-";
            if (date.Day < 10) res += "0" + date.Day.ToString();
            else res += date.Day.ToString();

            return res;
        }

        // Заброс данных в шаблон отчета Кадры-Экспресс.docx
        private void PutDataToReport(bool switched)
        {
            int icode = Convert.ToInt16(current_dg_code);

            report.LoadDocument("Кадры-Экспресс.docx");

            if (switched) mainTabControl.SelectedTabPageIndex = 3;

            Write2Report("month", Month2String(Convert.ToInt16(calc_month)));
            Write2Report("yy", calc_year);
            Write2Report("y1", calc_year[2].ToString());
            Write2Report("y2", calc_year[3].ToString());
            Write2Report("m1", calc_month[0].ToString());
            Write2Report("m2", calc_month[1].ToString());
            Write2Report("DG_name", current_dg_name);
            Write2Report("dgc1", current_dg_code[0].ToString());
            Write2Report("dgc2", current_dg_code[1].ToString());
            Write2Report("dgc3", current_dg_code[2].ToString());


            for (int i = 1; i < 42; i++)
            {
                    Write2Report( "pos" + i.ToString(), express[icode,i].ToString() );
            }

            Write2Report("ruk1_dol", viz_ruk1_dol.Text);
            Write2Report("ruk1_name", viz_ruk1_name.Text);
            Write2Report("ruk1_zvan", viz_ruk1_zvan.Text);

            Write2Report("ruk2_dol", viz_ruk2_dol.Text);
            Write2Report("ruk2_name", viz_ruk2_name.Text);
            Write2Report("ruk2_zvan", viz_ruk2_zvan.Text);

            Write2Report("isp1_name", viz_isp1_name.Text);
            Write2Report("isp1_zvan", viz_isp1_zvan.Text);
            Write2Report("script_date", script_date.Value.ToShortDateString());

            Write2Report("isp2_name", viz_isp2_name.Text);
            Write2Report("isp2_zvan", viz_isp2_zvan.Text);
            
        }

        // Сохранение данных ДГ в базу...
        private void SaveDGtoBase(string code, string date)
        {
            DateTime cdate = Convert.ToDateTime(date);
            string cmd = "";
            int form_code = Convert.ToInt16(code);

            // Проверяем есть ли такие данные в БД
            int res = DataProvider._getDataSQLs(sqlConnection,String.Format("SELECT COUNT(code_dg) AS CNT FROM Express_2012 WHERE (code_dg = {0}) AND (calc_date = '{1}-{2}-{3}')", code, cdate.Year, cdate.Month, cdate.Day));
            if (res > 0)
            {
                if (MessageBox.Show("Перезаписать существующие данные?", "Сохранение...", MessageBoxButtons.YesNo) != System.Windows.Forms.DialogResult.Yes) return;
                else
                {
                    cmd = "UPDATE Express_2012 SET ";
                    for (int i = 1; i < 42; i++)
                    {
                        if (i == 1) cmd += "pos_" + i.ToString() + " = " + express[form_code,i];
                        else cmd += ", pos_" + i.ToString() + " = " + express[form_code, i];
                    }
                    cmd += String.Format(" WHERE (code_dg = {0}) AND (calc_date = '{1}-{2}-{3}')", code, cdate.Year, cdate.Month, cdate.Day);
                    res = DataProvider._insDataSQL(sqlConnection,cmd);

                    //if (res == 1) MessageBox.Show("Данные успешно сохранены!");
                    //else MessageBox.Show("Ошибка при сохранении данных!");
                }
            }
            else
            {
                cmd = String.Format("INSERT INTO Express_2012 (code_dg, calc_date, pos_1, pos_2, pos_3, pos_4, pos_5, pos_6, pos_7, pos_8, pos_9, pos_10, pos_11, pos_12, pos_13, " +
                    "pos_14, pos_15, pos_16, pos_17, pos_18, pos_19, pos_20, pos_21, pos_22, pos_23, pos_24, pos_25, pos_26, pos_27, pos_28, pos_29, pos_30, pos_31, pos_32, pos_33, " +
                    "pos_34, pos_35, pos_36, pos_37, pos_38, pos_39, pos_40, pos_41) VALUES ('{0}','{1}-{2}-{3}',", code, cdate.Year, cdate.Month, cdate.Day);
                for (int i = 1; i < 42; i++)
                {
                    if (i == 1) cmd += express[form_code, i];
                    else cmd += ", " + express[form_code, i];
                }

                cmd += ")";

                res = DataProvider._insDataSQL(sqlConnection,cmd);

               // if (res == 1) MessageBox.Show("Данные успешно сохранены!");
               // else MessageBox.Show("Ошибка при сохранении данных!");
            }
        }

        // Загрузка основной формы
        private void MainForm_Load(object sender, EventArgs e)
        {
            Configuration cfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            word_report_path_text.Text = cfg.AppSettings.Settings["word_report_path_text"].Value;
            viz_ruk1_dol.Text = cfg.AppSettings.Settings["viz_ruk1_dol"].Value;
            viz_ruk1_dol.Text = cfg.AppSettings.Settings["viz_ruk1_zvan"].Value;
            viz_ruk1_name.Text = cfg.AppSettings.Settings["viz_ruk1_name"].Value;
            viz_ruk2_dol.Text = cfg.AppSettings.Settings["viz_ruk2_dol"].Value;
            viz_ruk2_zvan.Text = cfg.AppSettings.Settings["viz_ruk2_zvan"].Value;
            viz_isp1_zvan.Text = cfg.AppSettings.Settings["viz_isp1_zvan"].Value;
            viz_isp2_zvan.Text = cfg.AppSettings.Settings["viz_isp2_zvan"].Value;
            viz_ruk2_name.Text = cfg.AppSettings.Settings["viz_ruk2_name"].Value;
            viz_isp1_name.Text = cfg.AppSettings.Settings["viz_isp1_name"].Value;
            viz_isp2_name.Text = cfg.AppSettings.Settings["viz_isp2_name"].Value;
            calc_on_date.Value = Convert.ToDateTime(cfg.AppSettings.Settings["calc_on_date"].Value);
            script_date.Value = Convert.ToDateTime(cfg.AppSettings.Settings["script_date"].Value);
            reserv_date.Value = Convert.ToDateTime(cfg.AppSettings.Settings["reserv_date"].Value);

            try
            {
                // TODO: данная строка кода позволяет загрузить данные в таблицу "perechenDataSet.Express_Pergroup". При необходимости она может быть перемещена или удалена.
                this.express_PergroupTableAdapter.Fill(this.perechenDataSet.Express_Pergroup);
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к БД и загрузки перечня форм! Проверьте настройки подключения...");
                Close();
            }
            
            // Параметры соединения с БД во вкладке "настройки"
            sql_connection_edit.Text = EasyExpress.Properties.Settings.Default.IASConnection;
            kadry_connection_edit.Text = EasyExpress.Properties.Settings.Default.kadryConnection;
            kadryConnection = kadry_connection_edit.Text;
            sqlConnection = sql_connection_edit.Text;
                      
            // Меняем раскладку клавиатуры...
            //InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new CultureInfo("ru-RU"));
            report.LoadDocument("Кадры-Экспресс.docx");

            // Инициализируем расчетный массив
            InitAllExpress(express);

            // Проверяем доступные отчеты
            express_dates.Items.Clear();
            past_express.Items.Clear();

            // Заполняем таблицу с перечнем на панели
            try
            {
                DataTable dt = DataProvider._getDataSQL(sqlConnection, "SELECT distinct calc_date FROM Express_2012 ORDER BY calc_date desc");
                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string edate = Convert.ToDateTime(dt.Rows[i]["calc_date"]).ToShortDateString();
                        express_dates.Items.Add(edate);
                        past_express.Items.Add(edate);
                    }
                    express_dates.SelectedIndex = 0;
                    past_express.SelectedIndex = 1;
                }
            }
            catch
            {
                MessageBox.Show("Ошибка подключения к БД! Проверьте настройки подключения...");
                Close();
            }

            

            calc_on_date.Value = DateTime.Now;
            script_date.Value = DateTime.Now;
            res_date = reserv_date.Value;
        }

        // Расчет одной выбранной ДГ
        private void calc_one_formButton_Click(object sender, EventArgs e)
        {
            CalcDG(current_dg_code, current_sql_text);
        }

        // Выход
        private void exit_Button_Click(object sender, EventArgs e)
        {
            GC.Collect();
            Close();
        }

        // Обработка нажатия на клавишу мыши в таблице должностных групп...
        private void grid_MouseClick(object sender, MouseEventArgs e)
        {
            GridArea area = grid.HitTest(e.X, e.Y);

            if (area == GridArea.Cell)
            {
                GridEXRow row = grid.CurrentRow;

                TitleDG.Text = "Текущая должностная группа: " + row.Cells["KEY_OTCH"].Text;
                current_dg_code = row.Cells["KEY_OTCH"].Text;
                current_sql_text = row.Cells["TEXT_QRY"].Text;
                current_dg_name = row.Cells["NAME_FORM"].Text;
                FillEditPanel(current_dg_code, 0);
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        // Сохранение текущей должностной группы в БД
        private void save_DG_button_Click(object sender, EventArgs e)
        {
            report.SaveDocumentAs();
        }

        // Проверка логики текущей ДГ
        private void check_logic_dg_button_Click(object sender, EventArgs e)
        {
            if (CheckDG(current_dg_code) == false)
            {
                mainTabControl.SelectedTabPageIndex = 2;
                MessageBox.Show("Форма проверена, есть ошибки...");
            }
            else
            {
                MessageBox.Show("Форма проверена, ошибок не обнаружено!");
            }
        }

        private void расчитатьДолжностнуюГруппуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StatusBar.Panels[0].Text = String.Format("[{0}] - {1}", current_dg_code, current_dg_name);
            CalcDG(current_dg_code, current_sql_text);
            FillEditPanel(current_dg_code,1);
            RefreshPosFromEditPanel();  //?
            //PutDataToReport();
        }

        // Редактирование текущей ДГ
        private void редактироватьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (current_dg_code != "")
            {
                mainTabControl.SelectedTabPageIndex = 1;
                FillEditPanel(current_dg_code, 1);
                RefreshPosFromEditPanel(); //?
            }
        }

        // Расчет всех должностных групп
        private void calc_all_button_Click(object sender, EventArgs e)
        {
            DataRowCollection rc = perechenDataSet.Tables[0].Rows;

            for (int i = 0; i < rc.Count; i++)
            {
                string code = rc[i]["KEY_OTCH"].ToString();
                if (code != "300" && code != "301" && code != "900" && code != "901" && code != "399" && code != "829" && code != "879") // С 01.10.2016 - нет ОВО, СОБР, ОМОН, ЛРР
                {
                    StatusBar.Panels[0].Text = String.Format("[{0}] - {1}", code, rc[i]["NAME_FORM"].ToString());
                    CalcDG(code, rc[i]["TEXT_QRY"].ToString());
                    
                }
                Application.DoEvents();
            }

            StatusBar.Panels[0].Text = "Расчет окончен, не забудьте сохраниться...";
            save_all_express_button.Enabled = true;
            save_all_word_button.Enabled = true;
        }

        // кнопка сохранения всего отчета
        private void save_all_express_button_Click(object sender, EventArgs e)
        {
            SaveAllExpressToBase();
        }

        // Проверка всего отчета
        private void check_logic_all_button_Click(object sender, EventArgs e)
        {
            if (!CheckInterDG())
            {
                mainTabControl.SelectedTabPageIndex = 2;
                MessageBox.Show("Отчет проверен, есть ошибки...");
            }
            else MessageBox.Show("Отчет проверен, ошибок НЕТ !!!");
        }

        // Загрузка данных всего отчета
        private void load_all_express_button_Click(object sender, EventArgs e)
        {
            DateTime edate = Convert.ToDateTime(express_dates.SelectedItem);
            LoadAllExpressFromBase(edate, express); 
        }

        private void load_express_date_button_Click(object sender, EventArgs e)
        {
            DateTime edate = Convert.ToDateTime(express_dates.SelectedItem);
            LoadAllExpressFromBase(edate, express);
            save_all_word_button.Enabled = true;
        }

        // Сохранение текущей ДГ в окне редактирования
        private void button4_Click(object sender, EventArgs e)
        {
            RefreshPosFromEditPanel();
            SavePos2Express(current_dg_code);
            SaveDGtoBase(current_dg_code, System.DateTime.Now.ToShortDateString());
            MessageBox.Show("Данные успешно сохранены!");
        }

        // Проверка ДГ в форме редактиования
        private void button3_Click(object sender, EventArgs e)
        {
            if (!CheckDG(current_dg_code)) mainTabControl.SelectedTabPageIndex = 2;
        }

        // Сложение 2-х форм: [src] прибавляют к [dst]
        private void SumDG(int src, int dst)
        {
            for (int i = 1; i < 42; i++)
                express[dst, i] += express[src, i];
        }

        // Добавить в текущую должностную группу
        private void button5_Click(object sender, EventArgs e)
        {
            int dst_code = Convert.ToInt16(current_dg_code);
            
            if (edit_plus1.Text != "" && MessageBox.Show(String.Format("Добавить к данным [{0}] данные [{1}] ?", dst_code, edit_plus1.Text),"Сложение ДГ",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                int src_code = Convert.ToInt16(edit_plus1.Text);
                SumDG(src_code, dst_code);
            }
            if (edit_plus2.Text != ""  && MessageBox.Show(String.Format("Добавить к данным [{0}] данные [{1}] ?", dst_code, edit_plus2.Text),"Сложение ДГ",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                int src_code = Convert.ToInt16(edit_plus2.Text);
                SumDG(src_code, dst_code);
            }
            if (edit_plus3.Text != ""  && MessageBox.Show(String.Format("Добавить к данным [{0}] данные [{1}] ?", dst_code, edit_plus3.Text),"Сложение ДГ",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                int src_code = Convert.ToInt16(edit_plus3.Text);
                SumDG(src_code, dst_code);
            }
            
            FillEditPanel(current_dg_code, 1);
        }

        // Очистка отчета и удаление его из БД
        private void clear_all_express_button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("ВНИМАНИЕ! Отчет будет полностью удален из БД! Вы уверены?", "???", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                InitAllExpress(express);
                int res = DataProvider._insDataSQL(sqlConnection, "DELETE FROM Express_2012");
                MessageBox.Show("Отчет обнулен!\nБаза очищена!");
            }
        }

        private void print_inedit_button_Click(object sender, EventArgs e)
        {
            PutDataToReport(true);
        }


        #region Переходы в форме редактирования..
        private void edit_pos1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CalcNek(Convert.ToInt16(current_dg_code));
                edit_pos2.Focus();
            }
        }

        private void edit_pos2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CalcNek(Convert.ToInt16(current_dg_code));
                edit_pos3.Focus();
            }
        }

        private void edit_pos3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos4.Focus();
        }

        private void edit_pos4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos5.Focus();
        }

        private void edit_pos5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos6.Focus();
        }

        private void edit_pos6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos7.Focus();
        }

        private void edit_pos7_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos8.Focus();
        }

        private void edit_pos8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos9.Focus();
        }

        private void edit_pos9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos10.Focus();
        }

        private void edit_pos10_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos11.Focus();
        }

        private void edit_pos11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos12.Focus();
        }

        private void edit_pos12_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos13.Focus();
        }

        private void edit_pos13_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos14.Focus();
        }

        private void edit_pos14_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos15.Focus();
        }

        private void edit_pos15_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos16.Focus();
        }

        private void edit_pos16_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos17.Focus();
        }

        private void edit_pos17_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos18.Focus();
        }

        private void edit_pos18_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos19.Focus();
        }

        private void edit_pos19_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos20.Focus();
        }

        private void edit_pos20_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos21.Focus();
        }

        private void edit_pos21_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos22.Focus();
        }

        private void edit_pos22_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos23.Focus();
        }

        private void edit_pos23_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos24.Focus();
        }

        private void edit_pos24_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos25.Focus();
        }

        private void edit_pos25_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos26.Focus();
        }

        private void edit_pos26_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos27.Focus();
        }

        private void edit_pos27_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos28.Focus();
        }

        private void edit_pos28_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos29.Focus();
        }

        private void edit_pos29_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos30.Focus();
        }

        private void edit_pos30_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos31.Focus();
        }

        private void edit_pos31_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos32.Focus();
        }

        private void edit_pos32_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos33.Focus();
        }

        private void edit_pos33_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos34.Focus();
        }

        private void edit_pos34_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos35.Focus();
        }

        private void edit_pos35_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos36.Focus();
        }

        private void edit_pos36_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos37.Focus();
        }

        private void edit_pos37_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos38.Focus();
        }

        private void edit_pos38_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos39.Focus();
        }

        private void edit_pos39_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos40.Focus();
        }

        private void edit_pos40_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos41.Focus();
        }

        private void edit_pos41_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos1.Focus();
        }

        #endregion

        // Экспорт данных отчета в формат Статистика-К
        private void export_button_Click(object sender, EventArgs e)
        {
            // Имя файла отчета
            saveExportFileDialog.FileName = String.Format("621{0}{1}.rep", calc_year.Substring(2, 2), calc_month); 
            // Диалог сохранения
            System.Windows.Forms.DialogResult res = saveExportFileDialog.ShowDialog();
            if (res != System.Windows.Forms.DialogResult.Cancel && saveExportFileDialog.FileName != "")
            {
                System.IO.StreamWriter rep = new StreamWriter(saveExportFileDialog.FileName);

                // Заголовок файла
                rep.WriteLine(String.Format("/h {0}{1}621 \\h",calc_year.Substring(2, 2), calc_month));
                // Начало первого раздела
                rep.WriteLine("/d 01 \\d");
                // Все формы в разделе...
                for(int i = 0; i < count_repDG; i++)
                {
                   rep.Write(repDG[i]);
                   if ( i < count_repDG-1) rep.Write(" ");
                }
                rep.WriteLine("");
                // Пишем данные...
                for (int i = 1; i < 42; i++)
                {
                    for(int j = 0; j < count_repDG; j++)
                    {
                        int code = Convert.ToInt16(repDG[j]);
                        rep.Write(express[code, i]);
                        if (j < count_repDG-1) rep.Write(" ");
                    }
                    rep.WriteLine("");
                }
                // Закрываем                
                rep.WriteLine("\\d");
                rep.WriteLine("\\End");
                rep.Flush();
                rep.Close();
                MessageBox.Show("Экспорт завершен!");
            }

        }

        // Сохранение всего отчета в формате Word
        private void save_all_word_button_Click(object sender, EventArgs e)
        {
            DataRowCollection rc = perechenDataSet.Tables[0].Rows;

                for (int i = 0; i < rc.Count; i++)
                {
                    current_dg_code = rc[i]["KEY_OTCH"].ToString();
                    current_dg_name = rc[i]["NAME_FORM"].ToString();
                    if (current_dg_code != "300" && current_dg_code != "301" && current_dg_code != "900" && current_dg_code != "901")
                    {
                        string fileName = String.Format("{0}\\{1}.docx", word_report_path_text.Text, current_dg_code);
                        StatusBar.Panels[0].Text = String.Format("[{0}] - {1}", current_dg_code, rc[i]["NAME_FORM"].ToString());
                        if (current_dg_code != "")
                        {
                            FillEditPanel(current_dg_code, 1);
                            RefreshPosFromEditPanel(); //?
                            PutDataToReport(false);
                            report.SaveDocument(fileName, DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
                            mainTabControl.SelectedTabPageIndex = 0;
                        }
                        Application.DoEvents();
                    }
                }
                StatusBar.Panels[0].Text = "Сохранение завершилось...";
        }

        // Получение пути для сохранения ворд отчетов
        private void get_dir_button_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                word_report_path_text.Text = folderBrowserDialog.SelectedPath;
            }
        }

        // Сравнение текущего отчета с предыдущим периодом
        private void check_past_button_Click(object sender, EventArgs e)
        {
            int[,] p_express = new int[1000, 42]; // Массив предыдущего отчета
            InitAllExpress(p_express);
            DateTime edate = Convert.ToDateTime(past_express.SelectedItem);
            LoadAllExpressFromBase(edate, p_express);
            debug.Items.Clear();
            int err = 0;

            for (int i = 0; i < 1000; i++)
            {
                if (DG.Contains(i))
                {
                    for (int j = 0; j < 42; j++)
                    {
                      // Сравниваем позиции отчетов
                      // Не проверяем:
                      // [1] штат
                      // [2] замещение
                      // [10] - [13] нахождение в распоряжении
                        if (express[i, j] < p_express[i, j] && (j < 10 || j > 13) && j != 2 && j != 1)
                        {
                            debug.Items.Add(String.Format("В форме {0} в позиции {1} значение меньше предыдущего периода: {2} < {3}", i, j, express[i, j], p_express[i, j]));
                            err++;
                        }
                    }
                }
            }
            MessageBox.Show(String.Format("Отчеты сравнены, ошибок: {0}",err));
            mainTabControl.SelectedTabPageIndex = 2;
        }

        // Print the form
        private void prn_form_button_Click(object sender, EventArgs e)
        {
            report.Print();
        }

        private void save_settings_btn_Click(object sender, EventArgs e)
        {
            Configuration cfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            cfg.AppSettings.Settings["word_report_path_text"].Value = word_report_path_text.Text;
            cfg.AppSettings.Settings["viz_ruk1_dol"].Value = viz_ruk1_dol.Text;
            cfg.AppSettings.Settings["viz_ruk1_zvan"].Value = viz_ruk1_dol.Text;
            cfg.AppSettings.Settings["viz_ruk1_name"].Value = viz_ruk1_name.Text;
            cfg.AppSettings.Settings["viz_ruk2_dol"].Value = viz_ruk2_dol.Text;
            cfg.AppSettings.Settings["viz_ruk2_zvan"].Value = viz_ruk2_zvan.Text;
            cfg.AppSettings.Settings["viz_isp1_zvan"].Value = viz_isp1_zvan.Text;
            cfg.AppSettings.Settings["viz_isp2_zvan"].Value = viz_isp2_zvan.Text;
            cfg.AppSettings.Settings["viz_ruk2_name"].Value = viz_ruk2_name.Text;
            cfg.AppSettings.Settings["viz_isp1_name"].Value = viz_isp1_name.Text;
            cfg.AppSettings.Settings["viz_isp2_name"].Value = viz_isp2_name.Text;
            cfg.AppSettings.Settings["calc_on_date"].Value = calc_on_date.Value.ToShortDateString();
            cfg.AppSettings.Settings["script_date"].Value = script_date.Value.ToShortDateString();
            cfg.AppSettings.Settings["reserv_date"].Value = reserv_date.Value.ToShortDateString();

            cfg.Save();
            MessageBox.Show("Настройки сохранены!");
        }
    }
}

/// Хронология изменений:
/// 23.05.2014 - добавил сохранение всего отчета в Word
/// 15.02.2014 - должностная группа [365] заменена должностными группами [360] и [362]
/// 24.10.2015 - добавил сравнение с предыдущим периодом
/// 20.12.2022 - переписано в соответствии с новой формой (последняя редакция)