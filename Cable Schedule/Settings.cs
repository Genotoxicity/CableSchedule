using System;
using System.Windows.Forms;
using System.Xml;
using e3;
using e3lib;

namespace Script
{
    /// <summary>
    /// Класс - окошко настроек
    /// </summary>
    public partial class Settings : Form
    {
        private e3Application App;
        private string xmlName; // имя xml файла

        public Settings(string xmlname, e3Application InApp)
        {
            InitializeComponent();
            xmlName = xmlname;  // имя xml документа 
            GetValues();    // получение настроек на контролы
            if (InApp != null) { App = InApp; } else { E3connector e3c = new E3connector(); App = e3c.Connect(); };
        }

        private void btn_apply_Click(object sender, EventArgs e)
        {
            XmlDocument Settings = new XmlDocument(); // xml файл конфигурации
            try
            { 
                Settings.Load(xmlName); // загрузка xml
                bool empty_string = true;   // проверка на пустые строки
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "shapka", tb_shapka.Text);    // записываем значения в файл
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "line", tb_line.Text);
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "formatfirst", tb_formatfirst.Text);
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "formatnext", tb_formatnext.Text);
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "prefix", tb_prefix.Text);
                XmlWork.SetInnerText(Settings, "code", ud_code.Value.ToString());
                XmlWork.SetInnerText(Settings, "includingCode", ud_includingCode.Value.ToString());
                XmlWork.SetInnerText(Settings, "x", ud_x.Value.ToString());
                XmlWork.SetInnerText(Settings, "yshapka", ud_yshapka.Value.ToString());
                XmlWork.SetInnerText(Settings, "shapkaheight", ud_shapkaheight.Value.ToString());
                XmlWork.SetInnerText(Settings, "lineheight", ud_lineheight.Value.ToString());
                XmlWork.SetInnerText(Settings, "ymin", ud_ymin.Value.ToString());
                XmlWork.SetInnerText(Settings, "yminnext", ud_yminnext.Value.ToString());
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "scale", tb_scale.Text);
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "marka", tb_marka.Text);
                XmlWork.SetInnerText(Settings, "KJmarka", tb_KJmarka.Text);
                XmlWork.SetInnerText(Settings, "length_sp", tb_length_sp.Text);
                if (cb_Delete.Checked == true) { XmlWork.SetInnerText(Settings, "delete", "1"); } else { XmlWork.SetInnerText(Settings, "delete", "0"); }
                if (cb_zero.Checked == true) { XmlWork.SetInnerText(Settings, "zero", "1"); } else { XmlWork.SetInnerText(Settings, "zero", "0"); }
                if (cb_checkforwires.Checked == true) { XmlWork.SetInnerText(Settings, "checkforwires", "1"); } else { XmlWork.SetInnerText(Settings, "checkforwires", "0"); }
                if (cb_oldlength.Checked == true) { XmlWork.SetInnerText(Settings, "oldlength", "1"); } else { XmlWork.SetInnerText(Settings, "oldlength", "0"); }
                XmlWork.SetInnerText(Settings, "coefficient", ud_coefficient.Value.ToString());
                XmlWork.SetInnerText(Settings, "additional", ud_additional.Value.ToString());
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "altitude", tb_altitude.Text);
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "place", tb_place.Text);
                empty_string = empty_string && XmlWork.SetInnerText(Settings, "purpose", tb_purpose.Text);
                XmlWork.SetInnerText(Settings, "carry_from", ud_carry_from.Value.ToString());
                XmlWork.SetInnerText(Settings, "carry_to", ud_carry_to.Value.ToString());
                XmlWork.SetInnerText(Settings, "carry_purpose", ud_carry_purpose.Value.ToString());
                XmlWork.SetInnerText(Settings, "carry_install_place", ud_carry_install_place.Value.ToString());
                XmlWork.SetInnerText(Settings, "purpose_opt1", cmb_opt1.SelectedIndex.ToString());
                XmlWork.SetInnerText(Settings, "purpose_opt2", cmb_opt2.SelectedIndex.ToString());
                XmlWork.SetInnerText(Settings, "purpose_opt3", cmb_opt3.SelectedIndex.ToString());
                XmlWork.SetInnerText(Settings, "purpose_opt4", cmb_opt4.SelectedIndex.ToString());
                if (!empty_string)
                {
                    MessageBox.Show("Оставлены пустыми ключевые значения. Изменения не будут сохранены.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
                }
                else
                {
                    if (String.IsNullOrEmpty(tb_formatnext.Text))   // если это поле оставлено пустым, то используем данные из tb_formatfirst
                    {
                        XmlWork.SetInnerText(Settings, "formatnext", tb_formatfirst.Text);
                        tb_formatnext.Text = tb_formatfirst.Text;
                    }
                    Settings.Save(xmlName);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Файл конфигурации " + xmlName + " не найден или имеет неправильный формат.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
            }
        }

        /// <summary>
        /// Получение значений
        /// </summary>
        private void GetValues()    
        {
            XmlDocument Settings = new XmlDocument();   // xml-ка
            try
            {
                Settings.Load(xmlName); // загрузка
                tb_shapka.Text = XmlWork.GetInnerText(Settings, "shapka");
                tb_line.Text = XmlWork.GetInnerText(Settings, "line");
                tb_formatfirst.Text = XmlWork.GetInnerText(Settings, "formatfirst");
                tb_formatnext.Text = XmlWork.GetInnerText(Settings, "formatnext");
                tb_prefix.Text = XmlWork.GetInnerText(Settings, "prefix");
                ud_code.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "code"));
                ud_includingCode.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "includingCode"));
                ud_x.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "x"));
                ud_yshapka.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "yshapka"));
                ud_shapkaheight.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "shapkaheight"));
                ud_lineheight.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "lineheight"));
                ud_ymin.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "ymin"));
                ud_yminnext.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "yminnext"));
                tb_scale.Text = XmlWork.GetInnerText(Settings, "scale");
                tb_marka.Text = XmlWork.GetInnerText(Settings, "marka");
                tb_KJmarka.Text = XmlWork.GetInnerText(Settings, "KJmarka");
                tb_length_sp.Text = XmlWork.GetInnerText(Settings, "length_sp");
                ud_coefficient.Value = Convert.ToDecimal(XmlWork.GetInnerText(Settings, "coefficient"));
                ud_additional.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "additional"));
                if (XmlWork.GetInnerText(Settings, "delete") == "1") { cb_Delete.Checked = true; } else { cb_Delete.Checked = false; }
                if (XmlWork.GetInnerText(Settings, "zero") == "1") { cb_zero.Checked = true; } else { cb_zero.Checked = false; }
                if (XmlWork.GetInnerText(Settings, "checkforwires") == "1") { cb_checkforwires.Checked = true; } else { cb_checkforwires.Checked = false; }
                if (XmlWork.GetInnerText(Settings, "oldlength") == "1") { cb_oldlength.Checked = true; } else { cb_oldlength.Checked = false; }
                tb_altitude.Text = XmlWork.GetInnerText(Settings, "altitude");
                tb_place.Text = XmlWork.GetInnerText(Settings, "place");
                tb_purpose.Text = XmlWork.GetInnerText(Settings, "purpose");
                ud_carry_from.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_from"));
                ud_carry_to.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_to"));
                ud_carry_purpose.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_purpose"));
                ud_carry_install_place.Value = Convert.ToInt32(XmlWork.GetInnerText(Settings, "carry_install_place"));
                cmb_opt1.SelectedIndex = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt1"));
                cmb_opt2.SelectedIndex = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt2"));
                cmb_opt3.SelectedIndex = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt3"));
                cmb_opt4.SelectedIndex = Convert.ToInt32(XmlWork.GetInnerText(Settings, "purpose_opt4"));
            }
            catch (Exception)
            {
                MessageBox.Show("Файл конфигурации " + xmlName + " не найден или имеет неправильный формат.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
            }
        }

        private void bt_settings_exit_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void Settings_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            
        }

        /// <summary>
        /// чтение из БД в текстбоксы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tb_scale_DoubleClick(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            DBView attributes = new DBView(App, DBViewGetData.AttributeDefinition);
            if (attributes.Connected)
            {
                attributes.ShowDialog();
                if (!String.IsNullOrEmpty(attributes.Result)) tb.Text = attributes.Result;
            }
        }

        private void tb_formatfirst_DoubleClick(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            DBView symbol = new DBView(App, DBViewGetData.SymbolData);
            if (symbol.Connected)
            {
                symbol.ShowDialog();
                if (!String.IsNullOrEmpty(symbol.Result)) tb.Text = symbol.Result;
            }
        }

        private void ud_code_DoubleClick(object sender, EventArgs e)
        {
            NumericUpDown ud = (NumericUpDown)sender;
            DBView code = new DBView(App, DBViewGetData.SchematicTypes);
            if (code.Connected)
            {
                code.ShowDialog();
                if (!String.IsNullOrEmpty(code.Result)) ud.Value = Convert.ToInt32(code.Result);
            }
        }

    }
}
