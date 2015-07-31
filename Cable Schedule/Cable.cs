using System;
using System.Collections.Generic;
using e3;


namespace Script
{
    /// <summary>
    /// Класс, абстрагирующий кабель
    /// </summary>
    public class Cable
    {
        private int id; // id кабеля
        private string name;    // имя ( позиционное обозначение )
        private string type;    // тип кабеля ( марка, имя изделия в БД )
        private string from;    // откуда ( прибор на первом конце жилы )
        private string to;  // куда ( второй конец жилы )
        private string from_place;  // место размещения откуда
        private string to_place;    // место размещения куда
        private string purpose; // назначение
        private decimal length; // длина
        private bool traced;    // отрассирован
        private bool connected; // подключен
        private bool significant; // учитывать ли кабель
        private bool connectedSymbol;//подключен к символу

        private List<int> netSegList;   // список идентификаторов участков цепей, проходящих по листам типа план трасс
        e3Job Prj;  // интерфейс проекта
        e3Device Cab;   // интерфес устройства ( кабеля )
        e3Device Dev1;  // еще для двух устройств
        e3Device Dev2;
        e3Component Com;    // интерфейс компонента ( изделие в базе данных )
        e3Pin Pin;  // интерфейс вывода

        // интерфейсы чтения-записи
        public int Id
        {
            get { return this.id; }
        }

        public string Name
        {
            get { return this.name; }
        }

        public string Type
        {
            get { return this.type; }
        }

        public string From
        {
            get { return this.from; }
        }

        public string To
        {
            get { return this.to; }
        }

        public string From_place
        {
            get { return this.from_place; }
        }

        public string To_place
        {
            get { return this.to_place; }
        }

        public string Purpose
        {
            get { return this.purpose; }
        }

        public decimal Length
        {
            get { return this.length; }
            set { this.length = value; }
        }

        public bool Traced
        {
            get { return this.traced; }
            set { this.traced = value; }
        }

        public bool Connected
        {
            get { return this.connected; }
        }

        public bool ConnectedSymbol
        {
            get { return this.connectedSymbol; }
        }

        public bool Significant
        {
            get { return this.significant; }
        }

        public List<int> NetSegList
        {
            get { return this.netSegList; }
        }

        public Cable(e3Application App, int cabId, string length_sp, string place_att, string purpose_att, int purpose_opt1,  int purpose_opt2,  int purpose_opt3, int purpose_opt4, List<int> codeSheetIds, List<int> includingCodeSheetIds, bool zero)
        {
            Prj = App.CreateJobObject();    // получаем объект проекта
            Cab = Prj.CreateDeviceObject(); // объект кабеля ( устройства )
            Cab.SetId(cabId);   // устанавливаем устройство
            dynamic pinIds = new dynamic[] { };
            int corCnt = Cab.GetPinIds(ref pinIds);   // получаем жилы кабеля
            if (zero)   // если учитывать НЕ все кабеля
            {
                significant = IsSignificant(Prj, pinIds, corCnt, includingCodeSheetIds);    // учитывать или нет
                if (!significant) { Cab.SetAttributeValue(length_sp, "0"); return; }    // если нет, обнулить атрибут длины ( для скрипта спецификации ), и выход
            }
            else significant = true;
            Dev1 = Prj.CreateDeviceObject();    // еще два устройства
            Dev2 = Prj.CreateDeviceObject();
            Com = Prj.CreateComponentObject();  // объект компонент ( изделие в БД )
            Pin = Prj.CreatePinObject();
            Com.SetId(cabId);   // устанавливаем компонент
            id = cabId; // устанавливаем требуемые значения в свойства
            name = Cab.GetName();
            type = Com.GetName();   // тип кабеля - его имя в БД
            purpose = "";
            length = 0;
            traced = false;
            connected = false;
            connectedSymbol = false;
            ProcessEnds(pinIds, corCnt, place_att, purpose_att, purpose_opt1, purpose_opt2, purpose_opt3, purpose_opt4);
            netSegList = GetNetSegmentList(Prj, pinIds, corCnt, codeSheetIds);
            if (netSegList.Count > 0) traced = true;
        }

        /// <summary>
        /// Функция проверяет прохождение кабеля по листам СВП
        /// </summary>
        /// <param name="Prj"></param>
        /// <param name="pinIds"></param>
        /// <param name="corCnt"></param>
        /// <param name="includingCodeSheetIds"></param>
        /// <returns></returns>
        private bool IsSignificant(e3Job Prj, dynamic pinIds, int corCnt, List<int> includingCodeSheetIds)
        {
            e3Pin Pin = Prj.CreatePinObject();  // объект жилы кабеля
            dynamic netSegIds = new dynamic[] { };  // массив идентификаторов участков цепей
            dynamic sheetIds = new dynamic[]{}; // массив идентификаторов листов
            int netSegCnt, sheetCnt, sheetId, netSegId;
            List<int> tempSheet = new List<int>();   // список, используется для хранения уже обработанных листов
            List<int> tempNetSeg = new List<int>();   // список, используется для хранения уже обработанных участков цепей
            for (int i = 1; i <= corCnt; i++)   // обход жил кабеля
            {
                Pin.SetId(pinIds[i]);
                netSegCnt = Pin.GetNetSegmentIds(ref netSegIds);    // идентификаторы участков цепей, через которые проходит жила.
                for (int j = 1; j <= netSegCnt; j++)
                {
                    netSegId = netSegIds[j];
                    if (tempNetSeg.Contains(netSegId)) continue;    // если уже обработана, переход к другой
                    tempNetSeg.Add(netSegId);   // иначе добавляем жилу в список обработанных
                    sheetCnt = Prj.GetItemSheetIds(netSegId, out sheetIds); // список листов, по которым проходит жила
                    for (int k = 1; k <= sheetCnt; k++)
                    {
                        sheetId = sheetIds[k];
                        if (tempSheet.Contains(sheetId)) continue;  // если лист уже учтен - продолжаем
                        tempSheet.Add(sheetId);
                        if (includingCodeSheetIds.Contains(sheetId)) return true;   // если лист есть в списке листов с определенным типом, то значит кабель проходит по правильным листам
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Функция устанавливает, подключен ли кабель, откуда и куда подключен, место размещения откуда и куда
        /// </summary>
        /// <param name="pinIds"></param>
        /// <param name="corCnt"></param>
        /// <param name="place_att"></param>
        /// <param name="purpose_att"></param>
        /// <param name="purpose_opt1"></param>
        /// <param name="purpose_opt2"></param>
        /// <param name="purpose_opt3"></param>
        /// <param name="purpose_opt4"></param>
        private void ProcessEnds(dynamic pinIds, int corCnt, string place_att, string purpose_att, int purpose_opt1, int purpose_opt2, int purpose_opt3, int purpose_opt4)
        {
            int temp_purpose = 0;   // вариант вывода назначения
            for (int i = 1; i <= corCnt; i++)
            {
                Pin.SetId(pinIds[i]);
                int EndPinId1 = Pin.GetEndPinId(1); // вывод на первом конце жилы ( откуда )
                int EndPinId2 = Pin.GetEndPinId(2); // куда
                if (EndPinId1 == 0 || EndPinId2 == 0) continue;   // если оба конца жилы не равны нулю, значит кабель подключен, иначе ищем следующую жилу
                if (connected) continue; // если кабель еще не подключен, значит его не обрабатывали, иначе пропускаем

                Dev1.SetId(EndPinId1);  // от выводов к изделиям, откуда
                Dev2.SetId(EndPinId2);  // куда
                if (Dev1.GetName() == "")
                    connectedSymbol = true;

                if (connectedSymbol) continue; //если кабель подключен к символу то пропускаем

                from = Dev1.GetAssignment();    // получаем значения шкафов для приборов
                to = Dev2.GetAssignment();
                if (String.IsNullOrEmpty(from) && String.IsNullOrEmpty(to)) // в зависимости наличия/отсутствия шкафов выбирается вариант для вывода назначения
                {
                    temp_purpose = purpose_opt1;
                }
                if (!String.IsNullOrEmpty(from) && String.IsNullOrEmpty(to))
                {
                    temp_purpose = purpose_opt2;
                }
                if (String.IsNullOrEmpty(from) && !String.IsNullOrEmpty(to))
                {
                    temp_purpose = purpose_opt3;
                }
                if (!String.IsNullOrEmpty(from) && !String.IsNullOrEmpty(to))
                {
                    temp_purpose = purpose_opt4;
                }
                if (String.IsNullOrEmpty(from)) // если шкафа нет, то выводим имя прибора без тире 
                {
                   from = RemoveDash(Dev1.GetName());
                }
                from_place = Dev1.GetAttributeValue(place_att); // место установки откуда
                if (String.IsNullOrEmpty(from_place))
                {
                    if (Dev1.IsTerminal() == 1)   // если устройство - то у него нужных атрибутов может не быть, тогда их надо смотреть с его клеммного блока
                    {
                        Dev1.SetId(Dev1.GetTerminalBlockId());  // вот и смотрим
                        from_place = Dev1.GetAttributeValue(place_att);
                    }
                }
                if (String.IsNullOrEmpty(to))
                {
                    to = RemoveDash(Dev2.GetName());
                }
                to_place = Dev2.GetAttributeValue(place_att);
                if (String.IsNullOrEmpty(to_place))
                {
                    if (Dev2.IsTerminal() == 1)
                    {
                        Dev2.SetId(Dev2.GetTerminalBlockId());
                        to_place = Dev2.GetAttributeValue(place_att);
                    }
                }
                switch (temp_purpose)   // устанавливаем назначение, несколько строк объединяем в одну (тект может быть с переносами)
                {
                    case 0: purpose = Dev1.GetAttributeValue(purpose_att).Replace("\r\n", " "); break;
                    case 1: purpose = Dev2.GetAttributeValue(purpose_att).Replace("\r\n", " "); break;
                    default: purpose = ""; break;
                }
                connected = true;
            }
        }

        /// <summary>
        /// Функция получает участки цепей кабеля, лежащие на планах трасс
        /// </summary>
        /// <param name="Prj"></param>
        /// <param name="pinIds"></param>
        /// <param name="corCnt"></param>
        /// <param name="codeSheetIds"></param>
        /// <returns></returns>
        private List<int> GetNetSegmentList(e3Job Prj, dynamic pinIds, int corCnt, List<int> codeSheetIds)
        {
            List<int> netSegmentList = new List<int>(); // собственно список
            e3Pin Pin = Prj.CreatePinObject();
            dynamic netSegIds = new dynamic[] { };
            dynamic sheetIds = new dynamic[] { };
            int netSegCnt, sheetCnt;
            for (int i = 1; i <= corCnt; i++)   // обход жил кабеля
            {
                Pin.SetId(pinIds[i]);   // выбор конкретной жилы
                netSegCnt = Pin.GetNetSegmentIds(ref netSegIds);    // участки цепей, по которым проходит жила 
                for (int j = 1; j <= netSegCnt; j++)
                {
                    if (netSegmentList.Contains(netSegIds[j])) continue;    // если сегмент уже обработан, переход к другому
                    sheetCnt = Prj.GetItemSheetIds(netSegIds[j], out sheetIds);
                    for (int k = 1; k <= sheetCnt; k++) if (codeSheetIds.Contains(sheetIds[k])) netSegmentList.Add(netSegIds[j]);   // проверка что участок цепи находится на листе планом трасс и добавление участка к списку
                }
            }
            return netSegmentList;   
        }

        /// <summary>
        /// Удаляет тире, если оно первый символ
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string RemoveDash(string str)
        {
            if (str[0].Equals('-')) { return str.Substring(1); } else { return str; }
        }
    }
}
