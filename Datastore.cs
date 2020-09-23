using System;


public  class TESTForm
{
    public TESTForm()
    {
        InitializeComponent();
    }

    public class GetResultFormulas
    {
        public string maxflow;
        public string maxtemp;
    }

    private void simpleButton1_Click(object sender, EventArgs e)
    {
        WriteXML();
    }

    private void WriteXML()
    {
        var conditions = new GetResultFormulas();
        conditions.maxflow = txtRFCMaxF.Text;
        conditions.maxtemp = txtITMaxF.Text;

        var xs = new XmlSerializer(typeof(GetResultFormulas));
        using (var fs = new FileStream("TestXML", FileMode.Create))
        {
            xs.Serialize(fs, conditions);
        }
    }

    private void simpleButton2_Click(object sender, EventArgs e)
    {
        ReadXML();
    }

    private void ReadXML()
    {
        try
        {
            GetResultFormulas conditions;
            XmlSerializer xs = new XmlSerializer(typeof(GetResultFormulas));
            using (FileStream fs = new FileStream("TestXML", FileMode.Open))
            {
                conditions = xs.Deserialize(fs) as GetResultFormulas;
            }
            if (conditions != null)
            {
                textEdit2.Text = conditions.maxflow;
                textEdit4.Text = conditions.maxtemp;
            }
        }
        catch (Exception)
        {

            return;
        }
    }
}

ComboBox cmbFluid;
ComboBox cmbState;

TextBox txtFRSO;
TextBox txtIPMaxF;
TextBox txtIPNF;
ComboBox cmbUnit;
TextBox txtIPMinF;
GroupBox groupBox1;

GroupBox groupBox2;
TextBox txtOPMaxF;
TextBox txtOPMinF;
TextBox txtOPNF;

GroupBox groupBox3;
TextBox txtITMaxF;
TextBox txtITMinF;
TextBox txtITNF;

GroupBox groupBox4;
TextBox txtFRMaxF;
TextBox txtFRMinF;
TextBox txtFRNF;

GroupBox groupBox9;
GroupBox groupBox10;
GroupBox groupBox11;
TextBox txtSWMWNF;
TextBox txtVSHRNF;

GroupBox groupBox14;
TextBox txtVVNF;
TextBox txtVVMaxF;

TextBox txtVVMinF;

GroupBox groupBox13;
TextBox txtRFCNF;
TextBox txtRFCMaxF;
TextBox txtRFCMinF;

TextBox txtITSNF;
TextBox txtITSMaxF;

TextBox txtITSMinF;

GroupBox groupBox15;
TextBox txtAMM;

TextBox txtLPD;
TextBox txtCPress;
ComboBox cmbIPUnit2;
ComboBox cmbIPUnit1;
ComboBox cmbITUnit;
GroupBox groupBox17;
TextBox txtAPSNF;
TextBox txtAPSMaxF;
TextBox txtAPSMinF;

ComboBox cmbITSUnit;
GroupBox groupBox6;

ComboBox cmbVBFD;

GroupBox groupBox7;
ComboBox cmbTType;
ComboBox cmbTBU;
ComboBox cmbTC;

ComboBox comboBox1;
GroupBox groupBox8;
TextBox textBox2;

TextBox textBox1;
GroupBox groupBox16;
TextBox textBox4;
TextBox textBox3;

TextBox textBox5;

TextBox textBox6;

TextBox textBox9;
TextBox textBox8;
TextBox textBox7;
TextBox cmbTSize;

GroupBox groupBox5;
GroupBox groupBox18;
ComboBox comboBox6;

ComboBox comboBox5;

ComboBox comboBox4;
ComboBox comboBox3;
ComboBox comboBox2;
ComboBox comboBox7;
Button button3;
Button button4;
GroupBox groupBox12;
CheckBox checkBox1;

GroupBox groupBox19;
CheckedListBox checkedListBox1;
ComboBox comboBox8;
TextBox textBox10;
RichTextBox richTextBox1;