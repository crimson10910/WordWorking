namespace WordWorking;

using System.Security.Cryptography;
using System.Windows.Forms;
using WordsChanger;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using word = Microsoft.Office.Interop.Word;

public partial class MainForm : Form
{
    private WordHelper _helper;
    DateTime dateTime;

    public MainForm()
    {
        InitializeComponent();
        //      Необходимо для плавного скроллинга
        //vScrollBar1.Hide();
        //vScrollBar1.Value = this.VerticalScroll.Value;
        //vScrollBar1.Minimum = this.VerticalScroll.Minimum;
        //vScrollBar1.Maximum = this.VerticalScroll.Maximum;
        WTYPE.SelectedIndex = 0;
        SX.SelectedIndex = 0;
        EDUCATION.SelectedIndex = 0;
        //ED_DOP_TYPE.SelectedIndex = 0;
        
        ///comboBox5.SelectedIndex = 0;
        WW1.SelectedIndex = 0;
        WW5.SelectedIndex = 0;
        ///comboBox5.SelectedIndex = 0;
        //dateTimePicker37.Value = new DateTime(1);
        

        this.AutoScroll = true;
        this.HorizontalScroll.Enabled = false;
        if (FlowLayoutPanel1_ControlAdded != null)
        {
            this.ControlAdded += FlowLayoutPanel1_ControlAdded;
        }
        if (FlowLayoutPanel1_ControlRemoved != null)
        {
            this.ControlRemoved += FlowLayoutPanel1_ControlRemoved;
        }

        dateTime = DATE.Value;

    }

    private void buttonChange_Click(object sender, EventArgs e)
    {
        _helper = new WordHelper(@"Files\t-2_TEMP.doc");  // создаем объект класса с параметром в
                                                        // виде местоположения файла

        var items = new Dictionary<string, string>      // Создаем словарь, где ключами будут теги,
                                                        // а информация из объектов на форме
            {
             { "<ORGANIZATION>", ORGANIZATION.Text },
             { "<DATE>" , DATE.Text },
             { "<TNAM>" , TNAM.Text },
             { "<INN>" , INN.Text },
             { "<STRAH>" , STRAH.Text },
             { "<A>" , A.Text },
             { "<WCHAR>" , WCHAR.Text },
             { "<WTYPE>" , WTYPE.Text },
             { "<SX>" , SX.Text },
             // общие сведения
             { "<WNUM>" , WNUM.Text },
             { "<WDATE>" , WDATE.Value.ToString("dd.MM.yy") },
             { "<LASTNAME>" , LASTNAME.Text },
             { "<FIRTSNAME>" , FIRTSNAME.Text },
             { "<PATRONIC>" , PATRONIC.Text },
             { "<BDAY>" , BDAY.Text },
             { "<BPLACE>" , BPLACE.Text },
             { "<GRAJ>" , GRAJ.Text },
             { "<LANGS>" , LANGS.Text },
             { "<EDUCATION>" , EDUCATION.Text },
             { "<KEY1>" , KEY1.Text },
             { "<KEY2>" , KEY2.Text },
             { "<KEY3>" , KEY3.Text },
             { "<KEY4>" , KEY4.Text },
             { "<KEY6>" , KEY6.Text },
             { "<KEY7>" , KEY7.Text },
             { "<KEY8>" , KEY8.Text },
             { "<KEY9>" , KEY9.Text },


             { "<ED_PL1>" , ED_PL1.Text },
             { "<ED_NAME1>" , ED_NAME1.Text },
             { "<ED_S1>" , ED_S1.Text },
             { "<ED_NU1>" , ED_NU1.Text },
             { "<ED_YEAR1>" , ED_YEAR1.Text },
             { "<ED_QWAL_11>" , ED_QWAL_11.Text },
             { "<ED_QWAL_12>" , ED_QWAL_12.Text },

             { "<ED_PL2>" , ED_PL2.Text },
             { "<ED_NAME2>" , ED_NAME2.Text },
             { "<ED_S2>" , ED_S2.Text },
             { "<ED_NU2>" , ED_NU2.Text },
             { "<ED_YEAR2>" , ED_YEAR2.Text },
             { "<ED_QWAL_21>" , ED_QWAL_21.Text },
             { "<ED_QWAL_22>" , ED_QWAL_22.Text },
             { "<ED_DOP_TYPE>" , ED_DOP_TYPE.Text },

             { "<ED_PL21>" , ED_PL21.Text },
             { "<ED_PL22>" , ED_PL22.Text },
             //{ "<ED_PL23>" , ED_PL23.Text },
             { "<ED_N1>" , ED_N1.Text },
             { "<ED_N2>" , ED_N2.Text },
             { "<ED_Y1>" , ED_Y1.Text },
             { "<ED_Y2>" , ED_Y2.Text },
             //{ "<KEY91>" , KEY91.Text },
             { "<KEY10>" , KEY10.Text },
             { "<KEY11>" , KEY11.Text },
             { "<PROF_BASE>" , PROF_BASE.Text },
             { "<PROF_OTHER>" , PROF_OTHER.Text },

             { "<D2>" , SPLIT8.Value.ToString("dd")},
             { "<M2>" , SPLIT8.Value.ToString("MM") },
             { "<Y2>" , SPLIT8.Value.ToString("yy") },

             { "<D3>" , D3.Value.ToString() },
             { "<M3>" , M3.Value.ToString() },
             { "<Y3>" , Y3.Value.ToString() },

             { "<D4>" ,  D4.Value.ToString()},
             { "<M4>", M4.Value.ToString() },
             { "<Y4>" , Y4.Value.ToString() },
             { "<D5>" ,  D5.Value.ToString()},
             { "<M5>" , M5.Value.ToString() },
             { "<Y5>" , Y5.Value.ToString() },

             { "<D6>" ,  D6.Value.ToString()},
             { "<M6>" , M6.Value.ToString() },
             { "<Y6>" , Y6.Value.ToString() },

             { "<MARREGE_STATUS>" , MARREGE_STATUS.Text },
             { "<KEY12>" , KEY12.Text },

             { "<ST_ROD1>" , ST_ROD1.Text },
             { "<ST_ROD2>" , ST_ROD2.Text },
             { "<ST_ROD3>" , ST_ROD3.Text },
             { "<ST_ROD4>" , ST_ROD4.Text },
             { "<ST_ROD5>" , ST_ROD5.Text },
             { "<ST_ROD6>" , ST_ROD6.Text },
             { "<ST_FIO1>" , ST_FIO1.Text },
             { "<ST_FIO2>" , ST_FIO2.Text },
             { "<ST_FIO3>" , ST_FIO3.Text },
             { "<ST_FIO4>" , ST_FIO4.Text },
             { "<ST_FIO5>" , ST_FIO5.Text },
             { "<ST_FIO6>" , ST_FIO6.Text },
             { "<ST_Y1>" , ST_Y1.Text },
             { "<ST_Y2>" , ST_Y2.Text },
             { "<ST_Y3>" , ST_Y3.Text },
             { "<ST_Y4>" , ST_Y4.Text },
             { "<ST_Y5>" , ST_Y5.Text },
             { "<ST_Y6>" , ST_Y6.Text },

             { "<PASS>" , PASS.Text },
             { "<PD>" , PASSDATE.Value.ToString("dd") },
             { "<PM>" , PASSDATE.Value.ToString("MM") },
             { "<PY>" , PASSDATE.Value.ToString("yy") },
             { "<PORG1>" , PORG.Text },
             { "<PORG2>" , "" },
             { "<PORG3>" , "" },
             { "<PASINDEX>" , PASINDEX.Text },
             { "<PASPL>" , PASPL.Text },
             { "<REALINDEX>" , REALINDEX.Text },
             { "<REALPL>" , REALPL.Text },
             { "<RD>" , RDATW.Value.Day.ToString() },
             { "<RM>" , RDATW.Value.Month.ToString() },
             { "<RY>" , RDATW.Value.Year.ToString() },
             { "<PHONENUMBER>" , PHONENUMBER.Text },
             // II
             { "<WW1>" , WW1.Text },
             { "<WW2>" , WW2.Text },
             { "<WW3>" , WW3.Text },
             { "<WW4>" , WW4.Text },
             { "<WW5>" , WW5.Text },
             { "<WW6>" , WW6.Text },
             { "<WW7A>" , WW7A.Text },
             { "<WW7B>" , WW7B.Text },
             { "<WW8>" , WW8.Text },
             // III
             { "<WO1>" , !WO1.Checked? "" : WO1.Value.ToString("dd.MM.yyyy") },
             { "<WO2>" , !WO2.Checked? "" : WO2.Value.ToString("dd.MM.yyyy") },
             { "<WO3>" , !WO3.Checked? "" : WO3.Value.ToString("dd.MM.yyyy") },
             { "<WO4>" , !WO4.Checked? "" : WO4.Value.ToString("dd.MM.yyyy") },
             { "<WO5>" , !WO5.Checked? "" : WO5.Value.ToString("dd.MM.yyyy") },
             { "<WO6>" , !WO6.Checked? "" : WO6.Value.ToString("dd.MM.yyyy") },
             { "<WO7>" , !WO7.Checked? "" : WO7.Value.ToString("dd.MM.yyyy") },
             { "<WO8>" , !WO8.Checked? "" : WO8.Value.ToString("dd.MM.yyyy") },
             { "<WO9>" , !WO9.Checked? "" : WO9.Value.ToString("dd.MM.yyyy") },
             { "<WO10>" , !WO10.Checked? "" : WO10.Value.ToString("dd.MM.yyyy") },
             { "<WO11>" , !WO11.Checked? "" : WO11.Value.ToString("dd.MM.yyyy") },
             { "<WO12>" , !WO12.Checked? "" : WO12.Value.ToString("dd.MM.yyyy") },
             { "<WO13>" , !WO13.Checked? "" : WO13.Value.ToString("dd.MM.yyyy") },

             { "<WP1>" , WP1.Text },
             { "<WP2>" , WP2.Text },
             { "<WP3>" , WP3.Text },
             { "<WP4>" , WP4.Text },
             { "<WP5>" , WP5.Text },
             { "<WP6>" , WP6.Text },
             { "<WP7>" , WP7.Text },
             { "<WP8>" , WP8.Text },
             { "<WP9>" , WP9.Text },
             { "<WP10>" , WP10.Text },
             { "<WP11>" , WP11.Text },
             { "<WP12>" , WP12.Text },
             { "<WP13>" , WP13.Text },

             { "<WR1>" , WR1.Text },
             { "<WR2>" , WR2.Text },
             { "<WR3>" , WR3.Text },
             { "<WR4>" , WR4.Text },
             { "<WR5>" , WR5.Text },
             { "<WR6>" , WR6.Text },
             { "<WR7>" , WR7.Text },
             { "<WR8>" , WR8.Text },
             { "<WR9>" , WR9.Text },
             { "<WR10>" , WR10.Text },
             { "<WR11>" , WR11.Text },
             { "<WR12>" , WR12.Text },
             { "<WR13>" , WR13.Text },

             { "<WS1>" , WS1.Text },
             { "<WS2>" , WS2.Text },
             { "<WS3>" , WS3.Text },
             { "<WS4>" , WS4.Text },
             { "<WS5>" , WS5.Text },
             { "<WS6>" , WS6.Text },
             { "<WS7>" , WS7.Text },
             { "<WS8>" , WS8.Text },
             { "<WS9>" , WS9.Text },
             { "<WS10>" , WS10.Text },
             { "<WS11>" , WS11.Text },
             { "<WS12>" , WS12.Text },
             { "<WS13>" , WS13.Text },

             { "<WT1>" , WT1.Text },
             { "<WT2>" , WT2.Text },
             { "<WT3>" , WT3.Text },
             { "<WT4>" , WT4.Text },
             { "<WT5>" , WT5.Text },
             { "<WT6>" , WT6.Text },
             { "<WT7>" , WT7.Text },
             { "<WT8>" , WT8.Text },
             { "<WT9>" , WT9.Text },
             { "<WT10>" , WT10.Text },
             { "<WT11>" , WT11.Text },
             { "<WT12>" , WT12.Text },
             { "<WT13>" , WT13.Text },
             // IV
             { "<AB1>" , AB1.Text },
             { "<AB2>" , AB2.Text },
             { "<AB3>" , AB3.Text },
             { "<AB4>" , AB4.Text },
             { "<AB5>" , AB5.Text },
             { "<AB6>" , AB6.Text },

             { "<AC1>" , AC1.Text },
             { "<AC2>" , AC2.Text },
             { "<AC3>" , AC3.Text },
             { "<AC4>" , AC4.Text },
             { "<AC5>" , AC5.Text },
             { "<AC6>" , AC6.Text },

             { "<AD1>" , AD1.Text },
             { "<AD2>" , AD2.Text },
             { "<AD3>" , AD3.Text },
             { "<AD4>" , AD4.Text },
             { "<AD5>" , AD5.Text },
             { "<AD6>" , AD6.Text },

             { "<AE1>" , !AE1.Checked? "" : AE1.Value.ToString("dd.MM.yyyy") },
             { "<AE2>" , !AE2.Checked? "" : AE2.Value.ToString("dd.MM.yyyy") },
             { "<AE3>" , !AE3.Checked? "" : AE3.Value.ToString("dd.MM.yyyy") },
             { "<AE4>" , !AE4.Checked? "" : AE4.Value.ToString("dd.MM.yyyy") },
             { "<AE5>" , !AE5.Checked? "" : AE5.Value.ToString("dd.MM.yyyy") },
             { "<AE6>" , !AE6.Checked? "" : AE6.Value.ToString("dd.MM.yyyy") },

             { "<AF1>" , AF1.Text },
             { "<AF2>" , AF2.Text },
             { "<AF3>" , AF3.Text },
             { "<AF4>" , AF4.Text },
             { "<AF5>" , AF5.Text },
             { "<AF6>" , AF6.Text },

             //V
            
             { "<BA1>" , !BA1.Checked? "" : BA1.Value.ToString("dd.MM.yyyy") },
             { "<BA2>" , !BA2.Checked? "" : BA2.Value.ToString("dd.MM.yyyy") },
             { "<BA3>" , !BA3.Checked? "" : BA3.Value.ToString("dd.MM.yyyy") },
             { "<BA4>" , !BA4.Checked? "" : BA4.Value.ToString("dd.MM.yyyy") },
             { "<BA5>" , !BA5.Checked? "" : BA5.Value.ToString("dd.MM.yyyy") },
             { "<BA6>" , !BA6.Checked? "" : BA6.Value.ToString("dd.MM.yyyy") },

             { "<BB1>" , !BK1.Checked? "" : BK1.Value.ToString("dd.MM.yyyy") },
             { "<BB2>" , !BK2.Checked? "" : BK2.Value.ToString("dd.MM.yyyy") },
             { "<BB3>" , !BK3.Checked? "" : BK3.Value.ToString("dd.MM.yyyy")},
             { "<BB4>" , !BK4.Checked? "" : BK4.Value.ToString("dd.MM.yyyy") },
             { "<BB5>" , !BK5.Checked? "" : BK5.Value.ToString("dd.MM.yyyy") },
             { "<BB6>" , !BK6.Checked? "" : BK6.Value.ToString("dd.MM.yyyy") },

             { "<BC1>" , BC1.Text },
             { "<BC2>" , BC2.Text },
             { "<BC3>" , BC3.Text },
             { "<BC4>" , BC4.Text },
             { "<BC5>" , BC5.Text },
             { "<BC6>" , BC6.Text },

              { "<BD1>" , BD1.Text },
             { "<BD2>" , BD2.Text },
             { "<BD3>" , BD3.Text },
             { "<BD4>" , BD4.Text },
             { "<BD5>" , BD5.Text },
             { "<BD6>" , BD6.Text },

              { "<BE1>" , BE1.Text },
             { "<BE2>" , BE2.Text },
             { "<BE3>" , BE3.Text },
             { "<BE4>" , BE4.Text },
             { "<BE5>" , BE5.Text },
             { "<BE6>" , BE6.Text },

             { "<BF1>" , BF1.Text },
             { "<BF2>" , BF2.Text },
             { "<BF3>" , BF3.Text },
             { "<BF4>" , BF4.Text },
             { "<BF5>" , BF5.Text },
             { "<BF6>" , BF6.Text },

              { "<BG1>" , BG1.Text },
             { "<BG2>" , BG2.Text },
             { "<BG3>" , BG3.Text },
             { "<BG4>" , BG4.Text },
             { "<BG5>" , BG5.Text },
             { "<BG6>" , BG6.Text },

              { "<BH1>" , BH1.Text },
             { "<BH2>" , BH2.Text },
             { "<BH3>" , BH3.Text },
             { "<BH4>" , BH4.Text },
             { "<BH5>" , BH5.Text },
             { "<BH6>" , BH6.Text },

             //VI
             { "<CA1>" , CA1.Text },
             { "<CA2>" , CA2.Text },
             { "<CA3>" , CA3.Text },
             { "<CA4>" , CA4.Text },


             { "<CB1>" , CB1.Text },
             { "<CB2>" , CB2.Text },
             { "<CB3>" , CB3.Text },
             { "<CB4>" , CB4.Text },


             { "<CC1>" , CC1.Text },
             { "<CC2>" , CC2.Text },
             { "<CC3>" , CC3.Text },
             { "<CC4>" , CC4.Text },


              { "<CD1>" , CD1.Text },
             { "<CD2>" , CD2.Text },
             { "<CD3>" , CD3.Text },
             { "<CD4>" , CD4.Text },


              { "<CE1>" , CE1.Text },
             { "<CE2>" , CE2.Text },
             { "<CE3>" , CE3.Text },
             { "<CE4>" , CE4.Text },


             { "<CF1>" , CF1.Text },
             { "<CF2>" , CF2.Text },
             { "<CF3>" , CF3.Text },
             { "<CF4>" , CF4.Text },

             { "<CG1>" , CG1.Text },
             { "<CG2>" , CG2.Text },
             { "<CG3>" , CG3.Text },
             { "<CG4>" , CG4.Text },


             //VII
             { "<DA1>" , DA1.Text },
             { "<DA2>" , DA2.Text },
             { "<DA3>" , DA3.Text },
             { "<DA4>" , DA4.Text },

             { "<DB1>" , DB1.Text },
             { "<DB2>" , DB2.Text },
             { "<DB3>" , DB3.Text },
             { "<DB4>" , DB4.Text },


             { "<DC1>" , DC1.Text },
             { "<DC2>" , DC2.Text },
             { "<DC3>" , DC3.Text },
             { "<DC4>" , DC4.Text },


              { "<DD1>" , !DD1.Checked? "" : DD1.Value.ToString("dd.MM.yyyy") },
             { "<DD2>" , !DD1.Checked? "" : DD1.Value.ToString("dd.MM.yyyy") },
             { "<DD3>" , !DD1.Checked? "" : DD1.Value.ToString("dd.MM.yyyy") },
             { "<DD4>" , !DD1.Checked? "" : DD1.Value.ToString("dd.MM.yyyy") },

             //VII
             /*
              { "<EA1>" , EA1.Text },
             { "<EA2>" , EA2.Text },
             { "<EA3>" , EA3.Text },
             { "<EA4>" ,EA4.Text },
             { "<EA5>" , EA5.Text },
             { "<EA6>" , EA6.Text },
             { "<EA7>" , EA7.Text },
             { "<EA8>" , EA8.Text },
             { "<EA9>" , EA9.Text },
             { "<EA10>" , EA10.Text },
             { "<EA11>" , EA11.Text },
             { "<EA12>" , EA12.Text },
             { "<EA13>" , EA13.Text },

             { "<EB1>" , .Text },
             { "<EB2>" , .Text },
             { "<EB3>" , .Text },
             { "<EB4>" , .Text },
             { "<EB5>" , .Text },
             { "<EB6>" , .Text },
             { "<EB7>" , .Text },
             { "<EB8>" , .Text },
             { "<EB9>" , .Text },
             { "<EB10>" , .Text },
             { "<EB11>" , .Text },
             { "<EB12>" , .Text },
             { "<EB13>" , .Text },

             { "<EC1>" , .Text },
             { "<EC2>" , .Text },
             { "<EC3>" , .Text },
             { "<EC4>" , .Text },
             { "<EC5>" , .Text },
             { "<EC6>" , .Text },
             { "<EC7>" , .Text },
             { "<EC8>" , .Text },
             { "<EC9>" , .Text },
             { "<EC10>" , .Text },
             { "<EC11>" , .Text },
             { "<EC12>" , .Text },
             { "<EC13>" , .Text },

             { "<ED1>" , .Text },
             { "<ED2>" , .Text },
             { "<ED3>" , .Text },
             { "<ED4>" , .Text },
             { "<ED5>" , .Text },
             { "<ED6>" , .Text },
             { "<ED7>" , .Text },
             { "<ED8>" , .Text },
             { "<ED9>" , .Text },
             { "<ED10>" , .Text },
             { "<ED11>" , .Text },
             { "<ED12>" , .Text },
             { "<ED13>" , .Text },

             { "<EE1>" , .Text },
             { "<EE2>" , .Text },
             { "<EE3>" , .Text },
             { "<EE4>" , .Text },
             { "<EE5>" , .Text },
             { "<EE6>" , .Text },
             { "<EE7>" , .Text },
             { "<EE8>" , .Text },
             { "<EE9>" , .Text },
             { "<EE10>" , .Text },
             { "<EE11>" , .Text },
             { "<EE12>" , .Text },
             { "<EE13>" , .Text },

              { "<EF1>" , .Text },
             { "<EF2>" , .Text },
             { "<EF3>" , .Text },
             { "<EF4>" , .Text },
             { "<EF5>" , .Text },
             { "<EF6>" , .Text },
             { "<EF7>" , .Text },
             { "<EF8>" , .Text },
             { "<EF9>" , .Text },
             { "<EF10>" , .Text },
             { "<EF11>" , .Text },
             { "<EF12>" , .Text },
             { "<EF13>" , .Text },

              { "<EG1>" , .Text },
             { "<EG2>" , .Text },
             { "<EG3>" , .Text },
             { "<EG4>" , .Text },
             { "<EG5>" , .Text },
             { "<EG6>" , .Text },
             { "<EG7>" , .Text },
             { "<EG8>" , .Text },
             { "<EG9>" , .Text },
             { "<EG10>" , .Text },
             { "<EG11>" , .Text },
             { "<EG12>" , .Text },
             { "<EG13>" , .Text },*/

             //IX
             { "<FA1>" , FA1.Text },
             { "<FA2>" , FA2.Text },
             { "<FA3>" , FA3.Text },
             { "<FA4>" , FA4.Text },
             { "<FA5>" , FA5.Text },
             { "<FA6>" , FA6.Text },

             { "<FB1>" , FB1.Text },
             { "<FB2>" , FB2.Text },
             { "<FB3>" , FB3.Text },
             { "<FB4>" , FB4.Text },
             { "<FB5>" , FB5.Text },
             { "<FB6>" , FB6.Text },

             { "<FC1>" , !FC1.Checked? "" : FC1.Value.ToString("dd.MM.yyyy") },
             { "<FC2>" , !FC2.Checked? "" : FC2.Value.ToString("dd.MM.yyyy") },
             { "<FC3>" , !FC3.Checked? "" : FC3.Value.ToString("dd.MM.yyyy") },
             { "<FC4>" , !FC4.Checked? "" : FC4.Value.ToString("dd.MM.yyyy") },
             { "<FC5>" , !FC5.Checked? "" : FC5.Value.ToString("dd.MM.yyyy") },
             { "<FC6>" , !FC6.Checked? "" : FC6.Value.ToString("dd.MM.yyyy") },

              { "<FD1>" , FD1.Text },
             { "<FD2>" , FD2.Text },
             { "<FD3>" , FD3.Text },
             { "<FD4>" , FD4.Text },
             { "<FD5>" , FD5.Text },
             { "<FD6>" , FD6.Text },
             //X
             { "<G1>" , GGG.Text },
             { "<G2>" , GGG.Text },
             { "<G3>" , GGG.Text },
             { "<G4>" , GGG.Text },
             { "<G5>" , GGG.Text },

             { "<OUTCASE>" , OUTCASE.Text },              
             { "<OD>" , DOUT.Value.ToString("dd") },
             { "<OM>" , DOUT.Value.ToString("MMM") },
             { "<OY>" , DOUT.Value.ToString("yy") },
             { "<ONAMU>" , ONAMU.Text },
             { "<OOD>" , DPOUT.Value.ToString("dd") },
             { "<OOM>" , DPOUT.Value.ToString("MMM") },
             { "<OOY>" , DPOUT.Value.ToString("yy")},
             
            //{ "<PROF>" , textBox3.Text },
            // { "<DATE_FROM>", dateTimePicker1.Value.ToString("dd.MM.yyyy") },
        };
        _helper.Process(items, chkShowPreview.Checked); //вызываем метод из объекта класса с параметрами 
    }
    private void FlowLayoutPanel1_ControlRemoved(object sender, ControlEventArgs e)
    {
       // vScrollBar1.Minimum = this.VerticalScroll.Minimum;
    }
    private void FlowLayoutPanel1_ControlAdded(object sender, ControlEventArgs e)
    {
       // vScrollBar1.Maximum = this.VerticalScroll.Maximum;
    }
    private void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
    {
        //this.VerticalScroll.Value = vScrollBar1.Value;
    }
    private void label4_Click(object sender, EventArgs e)
    {

    }

    private void label6_Click(object sender, EventArgs e)
    {

    }

    private void label5_Click(object sender, EventArgs e)
    {

    }

    private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
    {

    }

    private void flowLayoutPanel1_Paint_1(object sender, PaintEventArgs e)
    {

    }

    private void MainForm_Load(object sender, EventArgs e)
    {

    }

    private void tableLayoutPanel5_Paint(object sender, PaintEventArgs e)
    {

    }

    private void tabPage3_Click(object sender, EventArgs e)
    {

    }

    private void tabPage1_Click(object sender, EventArgs e)
    {

    }

    private void textBox77_TextChanged(object sender, EventArgs e)
    {

    }

    private void textBox99_TextChanged(object sender, EventArgs e)
    {

    }

    private void textBox19_TextChanged(object sender, EventArgs e)
    {

    }


    private void DATE_MouseUp(object sender, MouseEventArgs e)
    {
        if (((DateTimePicker)sender).Checked)
        {
            ((DateTimePicker)sender).Checked = false;
            return;
        }
        else
        {
            ((DateTimePicker)sender).Checked = true;
            return;
        }
    }
}