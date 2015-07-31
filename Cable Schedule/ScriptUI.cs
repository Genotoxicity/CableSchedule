using System;
using System.Collections.Generic;
using System.Windows.Forms;
using e3;
using e3lib;
using System.Xml;
using System.Threading;

// скрипт выводит кабельный журнал
namespace Script
{
    public partial class ScriptUI : Form
    {
        

        public ScriptUI()
        {
            InitializeComponent();
            xmlName = Application.ExecutablePath;   // формируем имя файла настроек
            xmlName = xmlName.ToLower();
            xmlName = xmlName.Replace(".exe", ".xml");
            if (!GetSettings()) Environment.Exit(0);    // проверяем файл настроек на целостность
            e3c = new E3connector();
        }

        private E3connector e3c;    // объект для доступа к функциям библиотеки e3lib
        private e3Application App;
        private string xmlName; // имя xml файла конфигурации
        private string sym_shapka;  // имя символа шапки
        private string sym_line;    // имя символа линии (строки)
        private string format_first; // форматка первого листа
        private string format_next; // форматка последующих листов
        private string prefix;  // префикс имени листов
        private int code;   // код типа схемы "план трасс"
        private int includingCode;   // код типа схемы, на которой кабели не должны учитываться
        private int x;  // координата x размещения точки привязки символов на листе
        private int y_shapka;   // координата y размещения точки привязки символа шапки на листе 
        private int shapka_height;  // высота шапки
        private int line_height;    // высота линии
        private int y_min;  // y перехода на новый лист с первого листа
        private int y_min_next; // y перехода на новый лист с последующих листов
        private string scale_att;   // имя атрибута "масштаб листа"
        private string marka;   // имя атрибута "марка листа"
        private string KJmarka; // марка для листов кабельного журнала
        private string length_sp; // атрибут длины кабеля
        private decimal coefficient;    // коэффициент запаса
        private int additional; // добавлять к длине кабелей
        private bool delete;    // флаг удаления старых листов КЖ
        private bool zero;    // флаг использования кабелей только на СВП
        private bool checkForWires;    // флаг проверки на провода
        private bool oldlength;    // флаг использования длины в атрибуте
        private string altitude;    // имя атрибута "величина высотного перехода"
        private string place_att;   // имя атрибута "место установки"
        private string purpose_att; // имя атрибута "назначение"
        private int carry_from; // количество символов до переноса строки в поле "откуда"
        private int carry_to; // количество символов до переноса строки в поле "куда"
        private int carry_purpose; // количество символов до переноса строки в поле "назначение"
        private int carry_install_place; // количество символов до переноса строки в поле "место установки"
        private int purpose_opt1;   // опции для вывода "назначения" в зависимости от выбора пользователя
        private int purpose_opt2;
        private int purpose_opt3;
        private int purpose_opt4;
        private int threadCount;
        private Dictionary<int, Cable> Cables;  // коллекция с кабелями (id, кабель )
        private Object Lock = new Object(); // для залочивания важных участков кода от многопоточного доступа
        private List<int> codeSheetIds;   // для идентификаторов листов типа план трасс
        private List<int> includingCodeSheetIds;  // для идентификаторов листов типа СВП

        // интерфейс записи-чтения
        public string XmlName
        {
            get { return this.xmlName;}
        }

        private void btn_Do_Click(object sender, EventArgs e)   // кнопка расчитать
        {
            App = e3c.Connect();  // выбираем проект
            if (App == null)
            {
                MessageBox.Show("Нет выбранных проектов", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                return;
            }
            if (!GetSettings()) Environment.Exit(0);    // получаем настройки ( и првоеряем на целостность )
            DateTime start = e3c.ScriptStart(App, "Кабельный журнал");  // выводим время начала скрипта, переменная старт нужна будет для вычисления времени работы
            CheckForIllegalCrossThreadCalls = false;    // доступен к элементам управления из под другого потока
            e3Job Prj = App.CreateJobObject();  // создаем интерфейс проекта
            codeSheetIds = new List<int>();   // для идентификаторов листов типа план трасс
            includingCodeSheetIds = new List<int>();  // для идентификаторов листов типа СВП
            GetSheetLists(App);    // получение идентификаторов листов нужных типов
            Cables = new Dictionary<int, Cable>();   // коллекция идентификатор - кабель
            if (delete) SheetDelete(App);   // при необходимости удаляем старые КЖ листы
            if (checkForWires) if (!CheckForWires(App)) { e3c.Clear(ref App, ref Prj); return; }
            lb_Status.Text = "Получение кабелей";
            GetCables(App); // заполняем коллекцию
            if (!oldlength)
            {
                if (!CalculateLength(App)) { e3c.Clear(ref App, ref Prj); return; }  // вычисление длин кабелей, проверка валидности планов трасс
            }
            else { GetLengthFromAttribute(App); }
            pb_Progress.Value = 0;
            lb_Status.Text = "Поиск не подключенных кабелей";
            GetNotConnectedCables(Cables);  // поиск не подключенных кабелей
            lb_Status.Text = "Поиск не отрассированных кабелей";
            if (GetNotTracedCables(Cables)) // поиск не оттрассированных кабелей
            {
                List<Cable> Cab = new List<Cable>(Cables.Values);   // получаем список кабелей, для удобной сортировки
                lb_Status.Text = "Сортировка";
                Cab.Sort(new SpecialComparer());    // сортировка
                lb_Status.Text = "Вывод на листы";
                Output(App, Cab);   // вывод информации на листы
            }
            e3c.ScriptEnd(App, "Кабельный журнал", start);  // вывод времени окончания и времени работы
            Cables = null;
            lb_Status.Text = "Завершено";
            ni_script.ShowBalloonTip(5000, "Кабельный журнал", "Все задачи выполнены",ToolTipIcon.Info); // всплывающая подсказка в трее об окончании работы скрипта
            e3c.Clear(ref App, ref Prj);    // освобождаем ресурсы, чтобы E3 откликался на действия пользователя
        }

        /// <summary>
        /// Функция получения настроек, возвращаемое значение - успешность операции
        /// </summary>
        /// <returns></returns>
        private bool GetSettings()
        {
            XmlDocument Settings = new XmlDocument();   // объект xml-документа
          //  try
            {
                Settings.Load(xmlName); // загружаем файл
                sym_shapka = XmlWork.GetInnerText(Settings, "shapka");  // последовательно считываем значения
                sym_line = XmlWork.GetInnerText(Settings, "line");
                format_first = XmlWork.GetInnerText(Settings, "formatfirst");
                format_next = XmlWork.GetInnerText(Settings, "formatnext");
                prefix = XmlWork.GetInnerText(Settings, "prefix");
                code = Convert.ToInt32(XmlWork.GetInnerText(Settings, "code"));
                includingCode = Convert.ToInt32(XmlWork.GetInnerText(Settings, "includingCode"));
                x = Convert.ToInt32(XmlWork.GetInnerText(Settings, "x"));
                y_shapka = Convert.ToInt32(XmlWork.GetInnerText(Settings, "yshapka"));
                shapka_height = Convert.ToInt32(XmlWork.GetInnerText(Settings, "shapkaheight"));
                line_height = Convert.ToInt32(XmlWork.GetInnerText(Settings, "lineheight"));
                y_min = Convert.ToInt32(XmlWork.GetInnerText(Settings, "ymin"));
                y_min_next = Convert.ToInt32(XmlWork.GetInnerText(Settings, "yminnext"));
                scale_att = XmlWork.GetInnerText(Settings, "scale");
                marka = XmlWork.GetInnerText(Settings, "marka");
                KJmarka = XmlWork.GetInnerText(Settings, "KJmarka");
                length_sp = XmlWork.GetInnerText(Settings, "length_sp");
                string coeffString = XmlWork.GetInnerText(Settings, "coefficient");
                string separator = System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;
                coeffString = coeffString.Replace(".", separator);
                coeffString = coeffString.Replace(",", separator);
                coefficient = Convert.ToDecimal(coeffString);
                additional = Convert.ToInt32(XmlWork.GetInnerText(Settings, "additional"));
                if (XmlWork.GetInnerText(Settings, "delete") == "1") { delete = true; } else { delete = false; }
                if (XmlWork.GetInnerText(Settings, "zero") == "1") { zero = true; } else { zero = false; }
                if (XmlWork.GetInnerText(Settings, "checkforwires") == "1") { checkForWires = true; } else { checkForWires = false; }
                if (XmlWork.GetInnerText(Settings, "oldlength") == "1") { oldlength = true; } else { oldlength = false; }
                altitude = XmlWork.GetInnerText(Settings, "altitude");
                place_att = XmlWork.GetInnerText(Settings, "place");
                purpose_att = XmlWork.GetInnerText(Settings, "purpose");
                carry_from = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_from"));
                carry_to = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_to"));
                carry_purpose = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_purpose"));
                carry_install_place = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_install_place"));
                purpose_opt1 = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt1"));
                purpose_opt2 = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt2"));
                purpose_opt3 = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt3"));
                purpose_opt4 = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt4"));
                threadCount = Environment.ProcessorCount;
                return true;
            }
           /* catch (Exception e) // если что - то не так с файлом, выведем инфу
            {
                MessageBox.Show("Файл конфигурации " + xmlName + " не найден или имеет неправильный формат: " + e.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                return false;
            }*/
        }

        /// <summary>
        /// Функция получает списки листов по типам: план трасс и СВП
        /// </summary>
        /// <param name="App"></param>
        private void GetSheetLists(e3Application App)
        {
            e3Job Prj = App.CreateJobObject();
            e3Sheet Sheet = Prj.CreateSheetObject();
            bool findCode = false;  // флаг наличия планов трасс на проекте, длина кабелей рассчитывается именно на них
            bool findIncludingCode = false; // флаг наличия планов трасс на проекте, длина кабелей рассчитывается именно на них
            dynamic sheetIds = new dynamic[] { };
            int sheetCnt = Prj.GetSheetIds(ref sheetIds);   // количество и идентификаторы листов
            for (int i = 1; i <= sheetCnt; i++)
            {
                int sheetId = sheetIds[i];
                Sheet.SetId(sheetIds[i]);
                dynamic type = new dynamic[] { };
                int typeCnt = Sheet.GetSchematicTypes(ref type);    // получение типов листов
                if (typeCnt != 1) continue;  // пропускаем если лист не имеет определенный тип
                if (type[1] == code)    // если лист - план трасс
                {
                    findCode = true;
                    if (!codeSheetIds.Contains(sheetId)) codeSheetIds.Add(sheetId);
                }
                if (type[1] == includingCode)    // если лист - СВП
                {
                    findIncludingCode = true;
                    if (!includingCodeSheetIds.Contains(sheetId)) includingCodeSheetIds.Add(sheetId);
                }
            }
            if (!findCode || !findIncludingCode)    // выход из приложения, если каки х- то типов лисвто нет.
            {
                if (!findCode && findIncludingCode)  MessageBox.Show("Нет листов с планами трасс на проекте. Выход", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                if (findCode && !findIncludingCode) MessageBox.Show("Нет листов со схемами внешних проводок на проекте. Выход", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                if (!findCode && !findIncludingCode) MessageBox.Show("Нет листов со схемами внешних проводок и планами трасс на проекте. Выход", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                e3c.Clear(ref App, ref Prj);
                Environment.Exit(0);
            }
        }

        /// <summary>
        /// Проверка на провода
        /// </summary>
        /// <param name="App"></param>
        /// <returns></returns>
        private bool CheckForWires(e3Application App)
        {
            lb_Status.Text = "Проверка на провода";
            e3Job Prj = App.CreateJobObject();
            e3Device Dev = Prj.CreateDeviceObject();
            e3NetSegment NetSeg = Prj.CreateNetSegmentObject();
            e3Pin Pin = Prj.CreatePinObject();
            e3Sheet Sheet = Prj.CreateSheetObject();
            List<int> cableCoreIds = new List<int>();   // список с уже отбработанными жилами.
            bool findWire, showMessageBox = false;  // флаг найденного провода и флаг отображения мессаджбокса
            int netSegCnt, pinCnt, pinId;
            dynamic netSegIds = new dynamic[] { };  // идентификаторы сегментов цепи
            dynamic pinIds = new dynamic[] { }; // идентификаторы жил в кабеле
            string sheetName;   // имя листа
            foreach (int sheetId in codeSheetIds)   // обход листов с планами трасс
            {
                Sheet.SetId(sheetId);
                sheetName = Sheet.GetAttributeValue(marka).ToString() + "/" + Sheet.GetName().ToString();   // получение удобного имени
                findWire = false;   
                netSegCnt = Sheet.GetNetSegmentIds(ref netSegIds);  // получение участков цепи на листе
                for ( int i=1; i<=netSegCnt; i++)
                {
                    NetSeg.SetId(netSegIds[i]); // обход участков цепи
                    pinCnt = NetSeg.GetCoreIds(ref pinIds); // получение жил участка
                    for (int j=1; j<=pinCnt; j++)
                    {
                        pinId = pinIds[j];  // обход жил
                        if (cableCoreIds.Contains(pinId)) continue;
                        Dev.SetId(pinId);   // переход от жилы к кабелю/проводу
                        if (Dev.IsCable() == 1) { cableCoreIds.Add(pinId); } else { findWire = true; break; }   // если кабель - добавляем жилу в список, если провод - прерываем цикл перебора жил
                    }
                    if (findWire) break;    // если найден провод выход из цикла перебора участков цепи
                }
                if (findWire) { App.PutInfo(0, "На листе " + sheetName + " обнаружены провода"); showMessageBox = true; }
            }
            if (!showMessageBox) return true;
            if (MessageBox.Show(this, "Обнаружены провода на планах трасс!", "Продолжить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000) == DialogResult.Yes) { return true; } else { return false; };
        }

        /// <summary>
        /// Заполенение коллекции кабелями
        /// </summary>
        /// <param name="App"></param>
        /// <param name="Cables"></param>
        private void GetCables(e3Application App)
        {
            e3Job Prj = App.CreateJobObject();
            e3Device Dev = Prj.CreateDeviceObject();
            dynamic devIds = new dynamic[] { };
            int cabCnt = Prj.GetCableIds(ref devIds);   // получаем кабели проекта
            pb_Progress.Value = 0;
            pb_Progress.Maximum = cabCnt;
            int objCount = cabCnt / threadCount;    // примерное количество для обработки кабелей в разных потоках
            Thread[] t = new Thread[threadCount];
            for (int i = 1; i <= threadCount; i++)
            {
                int start, end;
                start = ((i - 1) * objCount) + 1;   // формируем начальный индекс элементов для потока
                if (i != threadCount) { end = i * objCount; } else { end = cabCnt; } // конечный
                t[i-1] = new Thread(delegate() { FillCables(App, devIds, start, end);}); // формируем поток
                t[i-1].Start(); // запускаем
            }
            e3c.WaitThreads(t,threadCount); // ждем завершения всех потоков
            Dictionary<int, Cable> temp = new Dictionary<int,Cable>();    // временная коллекция, для удаления неподходящих кабелей
            if (!zero) return;  // если учитывать все кабеля, то не надо удалять незначимые
            foreach (int index in Cables.Keys)
            {
                if (Cables[index].Significant) temp.Add(index, Cables[index]);   // добавление подходящих кабелей
            }
            Cables.Clear(); // очистка от всех значений
            Cables = temp;  // присваивание коллекции с правильными каблеями
            pb_Progress.Value = 0;
        }

        /// <summary>
        /// Функция заполнения коллекции кабелями в потоках
        /// </summary>
        /// <param name="App"></param>
        /// <param name="devIds"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        private void FillCables(e3Application App, dynamic devIds, int start, int end)
        {
            for (int i = start; i <= end; i++)
            {
                e3Job Prj = App.CreateJobObject();  // каждый раз формиурем проект и устройство чтобы не было тормозов
                e3Device Dev = Prj.CreateDeviceObject();
                Dev.SetId(devIds[i]);
                pb_Progress.Value++;
                if (Dev.IsCable() != 1) continue; // в список включаем только кабеля, не провода
                lock (Lock) { Cables.Add(devIds[i], new Cable(App, devIds[i], length_sp, place_att, purpose_att, purpose_opt1, purpose_opt2, purpose_opt3, purpose_opt4, codeSheetIds,  includingCodeSheetIds, zero)); } // формируем кабель со всеми нужными атрибутами
            } 
        }

        /// <summary>
        /// Функция считывает длины кабеля из атрибута
        /// </summary>
        /// <param name="App"></param>
        private void GetLengthFromAttribute(e3Application App)
        {
            lb_Status.Text = "Считывание длины из атрибута";
            string length;  // строка атрибута с длиной
            foreach (Cable Cab in Cables.Values)
            {
                e3Job Prj = App.CreateJobObject();
                e3Device Dev = Prj.CreateDeviceObject();
                Dev.SetId(Cab.Id);  // выбор кабеля
                length = Dev.GetAttributeValue(length_sp);  // считывание длины
                if (String.IsNullOrEmpty(length) || length.Equals("0"))
                {
                    Cab.Length = 0; // кабели с нулевой длиной не будут учитываться
                    continue;   // если пусто или ноль
                }
                try
                {
                    Cab.Length = Convert.ToDecimal(length);
                }
                catch
                {
                    Cab.Length = 0;
                    App.PutError(0, "У кабеля "+Dev.GetName().ToString()+" некорректное значение атрибута длины.");
                }
            }
            Dictionary<int, Cable> temp = new Dictionary<int, Cable>();    // временная коллекция, для удаления кабелей с нулевой длиной
            foreach (int index in Cables.Keys)
            {
                if (Cables[index].Length!=0) temp.Add(index, Cables[index]);   // добавление подходящих кабелей
            }
            Cables.Clear(); // очистка от всех значений
            Cables = temp;  // присваивание коллекции с правильными кабелями
        }

        /// <summary>
        /// Вычисление длины кабелей
        /// </summary>
        /// <param name="App"></param>
        /// <param name="Cables"></param>
        /// <returns></returns>
        private bool CalculateLength(e3Application App)
        {
            e3Job Prj = App.CreateJobObject();
            e3Sheet Sheet = Prj.CreateSheetObject();
            int scale = 1;  // масштаб по умолчанию
            foreach ( int codeSheetId in codeSheetIds)  // обход планов трасс
            {
                Sheet.SetId(codeSheetId);
                lbl_sheet.Text = "Лист " + Sheet.GetAttributeValue(marka).ToString() + "/" + Sheet.GetName();
                try // попытка получить масштаб
                {
                    int colon_pos = Sheet.GetAttributeValue(scale_att).LastIndexOf(":"); // получаем масштаб из надписи типа "1:500"
                    scale = Convert.ToInt32(Sheet.GetAttributeValue(scale_att).Substring(colon_pos + 1));
                }
                catch
                {
                    MessageBox.Show("Не задан масштаб у листа " + Sheet.GetAttributeValue(marka) + " " + Sheet.GetName() + ". Скрипт остановлен.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                    return false;
                }
                AddAltitude(App, codeSheetId);   // добавляем высотные переходы
                dynamic netsegIds = new dynamic[] { };
                int netsegCnt = Sheet.GetNetSegmentIds(ref netsegIds);  // количество и идентификаторы участков цепей
                lb_Status.Text = "Вычисление длины";
                pb_Progress.Maximum = netsegCnt;
                pb_Progress.Value = 0;
                int objCount = netsegCnt / threadCount; // так же разбиваем на потоки
                Thread[] t = new Thread[threadCount];
                for (int j = 1; j <= threadCount; j++)
                {
                    int start, end;
                    start = ((j - 1) * objCount) + 1;
                    if (j != threadCount) { end = j * objCount; } else { end = netsegCnt; }
                    t[j - 1] = new Thread(delegate() { FillNetSeg(App,netsegIds, scale, start, end); });
                    t[j - 1].Start();
                }
                e3c.WaitThreads(t,threadCount);
            }
            lbl_sheet.Text = String.Empty;
            lb_Status.Text = "Корректировка длины";
            pb_Progress.Value = 0;
            pb_Progress.Maximum = Cables.Count;
            foreach (Cable cab in Cables.Values)    // после вычислений на листах, финальные штрихи
            {
                cab.Length *= coefficient / 1000; // умножаем на коэффициент запаса и переводим в метры
                if (cab.Length % 1 != 0) cab.Length = (int)cab.Length + 1;  // округляем в большую сторону, если нужно
                cab.Length += additional;   // добавляем дополнительные метры
                if (cab.Connected && cab.Traced)   // только для подключенных и отрассированных кабелей
                {
                    e3Device Cab = Prj.CreateDeviceObject();
                    Cab.SetId(cab.Id);  // записываем длину в атрибут
                    Cab.SetAttributeValue(length_sp, cab.Length.ToString());
                }
                pb_Progress.Value++;
            }
            return true;
        }

        /// <summary>
        /// Функция подсчитывает длину кабеля на участках цепи и суммирует
        /// </summary>
        /// <param name="App"></param>
        /// <param name="netsegIds"></param>
        /// <param name="scale"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        private void FillNetSeg(e3Application App, dynamic netsegIds, int scale, int start, int end)
        {
            int netSegId;
            for (int i = start; i <= end; i++)
            {
                e3Job Prj = App.CreateJobObject();  // для быстродействия
                e3NetSegment NetSeg = Prj.CreateNetSegmentObject(); // для подсчета длины участка цепи
                netSegId = netsegIds[i];
                NetSeg.SetId(netSegId);
                foreach (Cable Cab in Cables.Values)    // проходжим по всем кабелям
                {
                    if (!Cab.NetSegList.Contains(netSegId)) continue;   // смотрим, проходит ли кабель через данный участок цепи
                    lock (Lock) Cab.Length += (decimal)NetSeg.GetSchemaLength() * scale;   // если проходит - плюсуем длину участка, умноженную на масштаб
                }
                pb_Progress.Value++;
            }
        }
   
        /// <summary>
        /// Добавление высотного перехода
        /// </summary>
        /// <param name="App"></param>
        /// <param name="sheetId"></param>
        /// <param name="Cables"></param>
        private void AddAltitude(e3Application App, int sheetId)
        {
            e3Job Prj = App.CreateJobObject();
            e3Sheet Sheet = Prj.CreateSheetObject();
            Sheet.SetId(sheetId);
            dynamic symIds = new dynamic[] { }; // массив идентификаторов симвлов на листе
            lb_Status.Text = "Добавление высотных переходов";
            int symCnt = Sheet.GetSymbolIds(ref symIds);    // символы на листе
            pb_Progress.Maximum = symCnt;
            pb_Progress.Value = 0;
            int objCount = symCnt / threadCount;
            Thread[] t = new Thread[threadCount];
            for (int i = 1; i <= threadCount; i++)
            {
                int start, end;
                start = ((i - 1) * objCount) + 1;
                if (i != threadCount) { end = i * objCount; } else { end = symCnt; }
                t[i - 1] = new Thread(delegate() { SymAltitude(App, symIds, start, end, sheetId); });    // вот тут и добавляем
                t[i - 1].Start();
            }
            e3c.WaitThreads(t, threadCount);
        }

        /// <summary>
        /// Вычисление высотного перехода у символа и добавление к длине кабеля
        /// </summary>
        /// <param name="App"></param>
        /// <param name="symIds"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="sheetId"></param>
        private void SymAltitude(e3Application App, dynamic symIds, int start, int end, int sheetId)
        { 
            e3Job Prj = App.CreateJobObject();
            e3Pin Pin = Prj.CreatePinObject();
            e3Symbol Sym = Prj.CreateSymbolObject();
            e3Sheet Sheet = Prj.CreateSheetObject();
            int netSegId, addToLength=0;
            string temp;
            dynamic pinIds = new dynamic[] { }; // идентификаторы жил кабеля
            dynamic netsegIds = new dynamic[] { };  // идентификаторы участков цепи
            for (int i = start; i <= end; i++)
            {
                Sym.SetId(symIds[i]);
                pb_Progress.Value++;
                temp = Sym.GetAttributeValue(altitude); // строковое значение атрибута высотного перехода
                if (String.IsNullOrEmpty(temp)) continue;   // если оно пустое, то пропускаем этот символ
                try   // попытка получить значение высотного перехода
                {
                    addToLength = Convert.ToInt32(Sym.GetAttributeValue(altitude));
                }
                catch (Exception)   // иначе
                {
                    string sheetName = Sheet.GetAttributeValue(marka).ToString() + "/" + Sheet.GetName().ToString();    // имя листа
                    App.PutInfo(0, "Некорректное значение высотного перехода (" + temp + ") у символа "+Sym.GetName().ToString()+" на листе "+sheetName);   // сообщение о неккоректном значении высотного перехода
                    continue;
                }
                int pinCnt = Sym.GetPinIds(ref pinIds); // доступ к кабелям можно получить через выводы символа, и участки цепи
                for (int j = 1; j <= pinCnt; j++)
                {
                    Pin.SetId(pinIds[j]);
                    int netsegCnt = Pin.GetNetSegmentIds(ref netsegIds);    // участки цепи подключенные к выводу
                    for (int k = 1; k <= netsegCnt; k++)
                    {
                        netSegId = netsegIds[k];
                        foreach (Cable Cab in Cables.Values)    // добавление высотного перехода. к длине тех кабелей, которые проходят через участок цепи с символом
                        {
                            if (!Cab.NetSegList.Contains(netSegId)) continue;
                            lock (Lock) Cab.Length += addToLength;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получение не подключенных кабелей
        /// </summary>
        /// <param name="Cables"></param>
        private void GetNotConnectedCables(Dictionary<int, Cable> Cables)
        {
            int notconnectedCount = 0; // количество не подключенных кабелей
            string notconnectedString = "";     // названия не подключенных кабелей
            int connectedSymbolCount = 0;
            string connectedSymbolString = "";
            foreach (Cable Cab in Cables.Values)
            {
                if (!Cab.Connected && !Cab.ConnectedSymbol)
                {
                    notconnectedCount++;
                    notconnectedString += (Cab.Name + "|");
                }
                else
                    if (!Cab.Connected && Cab.ConnectedSymbol)
                    {
                        connectedSymbolCount ++;
                        connectedSymbolString += (Cab.Name + "|");
                    }
            }
            pb_Progress.Maximum = Cables.Count - notconnectedCount;
            string message = "";
            if (notconnectedCount > 0 && connectedSymbolCount > 0)
                message = "На проекте обнаружены неподключенные кабеля(" + notconnectedCount.ToString() + "):\r\n" + notconnectedString + "\r\n и кабеля подключенные к символам(" + connectedSymbolCount.ToString() + "):\r\n" + connectedSymbolString;
            else
                if (notconnectedCount > 0)
                    message = "На проекте обнаружены неподключенные кабеля(" + notconnectedCount.ToString() + "):\r\n" + notconnectedString;
                else
                    if (connectedSymbolCount>0)
                        message = "На проекте обнаружены кабеля подключенные к символам(" + connectedSymbolCount.ToString() + "):\r\n" + connectedSymbolString;
            if (message!="") MessageBox.Show(message, "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
        }
        
        /// <summary>
        /// Получение неотрассированных кабелей, возвращаемое значение - выбор пользователя, продолжать или нет
        /// </summary>
        /// <param name="Cables"></param>
        /// <returns></returns>
        private bool GetNotTracedCables(Dictionary<int, Cable> Cables)
        {
            int nottracedCount = 0; // количество
            string nottracedString = "";    // названия
            foreach (Cable Cab in Cables.Values)
            {
                if (!Cab.Traced && Cab.Connected)
                {
                    nottracedCount++;
                    nottracedString += (Cab.Name + "|");
                }
            }
            pb_Progress.Maximum -= nottracedCount;
            if (nottracedCount > 0) // MessageBox с кнопками выбора
            {
                if (MessageBox.Show(this, "На проекте обнаружены неотрассированые кабеля(" + nottracedCount.ToString() + "):\r\n" + nottracedString, "Продолжить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000) == DialogResult.Yes) {return true;} else {return false;};
            }
            return true;
        }

        /// <summary>
        /// Вывод информации на листы
        /// </summary>
        /// <param name="App"></param>
        /// <param name="Cables"></param>
        private void Output(e3Application App, List<Cable> Cables)
        {
            e3Job Prj = App.CreateJobObject();
            e3Symbol Sym = Prj.CreateSymbolObject();
            int count = 0;  // количество листов
            string sheetName = prefix+count.ToString(); // имя листа
            int sheetId = 0;    // id листа
            double y = y_shapka-shapka_height;  // координата y для символов строк
            int min = y_min;
            e3Sheet Sheet = GetSheet(App, ref count, ref y, ref sheetId, format_first);  // создание листа с шапкой
            foreach (Cable Cab in Cables)
            {
                e3Job CabPrj = App.CreateJobObject();   // новый интерфейс для каждого листа, для быстродействия
                Sym = CabPrj.CreateSymbolObject();
                if (!Cab.Traced || !Cab.Connected) continue;    // если кабель не отрассирован или не подключен - не выводить
                StringWrapper From_place = new StringWrapper(Cab.From_place, carry_from);   // класс для переноса строк
                StringWrapper To_place = new StringWrapper(Cab.To_place, carry_to);
                StringWrapper Purpose = new StringWrapper(Cab.Purpose, carry_purpose);
                StringWrapper Install_place = new StringWrapper(Cab.From, carry_install_place);
                From_place.Wrap();  // функция получения части строки, влезающей в рамки
                To_place.Wrap();
                Purpose.Wrap();
                Install_place.Wrap();
                int SymId = Sym.Load(sym_line, null);   // загрузка символа строки
                SymId = Sym.Place(sheetId, x, y, null); // расположение на листе
                Sym.SetAttributeValue("column1", Cab.Name); // последовательно записываем значения
                if (Install_place.Write)    // Поле write указывает, писать текст или нет ( если пустой, или уже написан )
                {
                    Sym.SetAttributeValue("column10", Install_place.Text);  // Поле Text - текст для вывода
                    if (!Install_place.Resume) Install_place.Write = false; // Поле Resume определяет остался ли еще неперенесенный текст
                }
                if (From_place.Write)
                {
                    Sym.SetAttributeValue("column5", From_place.Text);
                    if (!From_place.Resume) From_place.Write = false;
                }
                Sym.SetAttributeValue("column11", Cab.To);
                if (To_place.Write)
                {
                    Sym.SetAttributeValue("column6", To_place.Text);
                    if (!To_place.Resume) To_place.Write = false;
                }
                Sym.SetAttributeValue("column2", Cab.Type);
                double length = (int)Cab.Length;
                Sym.SetAttributeValue("column7",length.ToString());
                if (Purpose.Write)
                {
                    Sym.SetAttributeValue("column0", Purpose.Text);
                    if (!Purpose.Resume) Purpose.Write = false;
                }
                y -= line_height;
                if (y < min) 
                {
                    Sheet.SetAttributeValue(marka, KJmarka);
                    CreateLine(Prj, sheetId, y);
                    Sheet = GetSheet(App, ref count, ref y, ref sheetId, format_next);
                    min = y_min_next; // переход на новый лист
                }
                while (From_place.Resume || To_place.Resume || Purpose.Resume || Install_place.Resume)  // пока все не выведем
                {
                    if (From_place.Resume) From_place.Wrap();
                    if (To_place.Resume) To_place.Wrap();
                    if (Purpose.Resume) Purpose.Wrap();
                    if (Install_place.Resume) Install_place.Wrap();
                    SymId = Sym.Load(sym_line, null);
                    SymId = Sym.Place(sheetId, x, y, null);
                    if (Install_place.Write)
                    {
                        Sym.SetAttributeValue("column10", Install_place.Text);
                        if (!Install_place.Resume) Install_place.Write = false;
                    }
                    if (From_place.Write)
                    {
                        Sym.SetAttributeValue("column5", From_place.Text);
                        if (!From_place.Resume) From_place.Write = false;
                    }
                    if (To_place.Write)
                    {
                        Sym.SetAttributeValue("column6", To_place.Text);
                        if (!To_place.Resume) To_place.Write = false;
                    }
                    if (Purpose.Write)
                    {
                        Sym.SetAttributeValue("column0", Purpose.Text);
                        if (!Purpose.Resume) Purpose.Write = false;
                    }
                    y -= line_height;
                    if (y < min)
                    {
                        Sheet.SetAttributeValue(marka, KJmarka);
                        CreateLine(Prj, sheetId, y);
                        Sheet = GetSheet(App, ref count, ref y, ref sheetId, format_next);
                        min = y_min_next;
                    }
                }
                pb_Progress.Value++;
            }
            CreateLine(Prj, sheetId, y);
            Prj = null;  // закрываем интерфейс для отображения листов     
        }

        /// <summary>
        /// Функция, создающая новый лист
        /// </summary>
        /// <param name="App"></param>
        /// <param name="count"></param>
        /// <param name="y"></param>
        /// <param name="sheetId"></param>
        /// <param name="format"></param>
        private e3Sheet GetSheet(e3Application App, ref int count, ref double y, ref int sheetId, string format)
        {
            e3Job Prj = App.CreateJobObject();  // создание интерфейса проекта
            e3Sheet Sheet = Prj.CreateSheetObject();
            e3Symbol Sym = Prj.CreateSymbolObject();
            count++;
            string sheetName = prefix + count.ToString();
            sheetId = Sheet.Create(0, sheetName, format, sheetId, 0); // создаем лист
            int SymId = Sym.Load(sym_shapka, null); // загружаем и размещаем символ шапки
            SymId = Sym.Place(sheetId, x, y_shapka, null);
            y = y_shapka - shapka_height;
            return Sheet;
        }

        private void CreateLine(e3Job Prj, int sheetId, double y)
        {
            e3Graph graph = Prj.CreateGraphObject();
            graph.CreateLine(sheetId, x, y, x - 395, y);
            graph.SetLineWidth(0.5d);
        }

        /// <summary>
        /// Функция, удаляющая листы КЖ
        /// </summary>
        /// <param name="App"></param>
        private void SheetDelete(e3Application App)
        {
            lb_Status.Text = "Удаление старых листов";
            e3Job Prj = App.CreateJobObject();
            e3Sheet Sheet = Prj.CreateSheetObject();
            dynamic sheetIds = new dynamic[] { };
            int sheetCnt = Prj.GetSheetIds(ref sheetIds);
            dynamic symIds = new dynamic[] { };
            for (int i = 1; i <= sheetCnt; i++) // перебираем все листы
            {
                Sheet.SetId(sheetIds[i]);
                string format = Sheet.GetFormat();
                if (format == format_first || format == format_next)    // если они нужного формата,
                {
                    int symCnt = Sheet.GetSymbolIds(ref symIds);
                    for (int j = 1; j <= symCnt; j++)
                    {
                        e3Job SymPrj = App.CreateJobObject();
                        e3Symbol Sym = SymPrj.CreateSymbolObject();
                        Sym.SetId(symIds[j]);
                        string type = Sym.GetSymbolTypeName();
                        if (type == sym_shapka || type == sym_line) // и на них есть символы нужного типа,
                        {
                            Sheet.Delete(); // то удаляем
                            break;
                        }
                    }
                }
            }
            
        }

        /// <summary>
        /// Вызываем окошко настроек
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Settings SettingsForm = new Settings(xmlName, App);
            SettingsForm.ShowDialog();
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ScriptUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            ni_script.Dispose();
            this.Dispose();
        }

        private void btn_script_exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ni_script_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ni_script.ShowBalloonTip(2000, "Кабельный журнал", "   THANK YOU MARIO!\r\n\r\nBUT OUR PRINCESS IS IN\r\n ANOTHER CASTLE!", ToolTipIcon.Info); 
        }

    }

    /// <summary>
    /// Класс, сравнивающий кабеля ( по именам )
    /// </summary>
    public class SpecialComparer : IComparer<Cable>
    {
        private static Comparator cmp = new Comparator();   // класс, сравнивающий строки (e3lib)
        
        public int Compare(Cable a, Cable b)
        {
            return cmp.Compare(a.Name,b.Name);
        }
    }

}
