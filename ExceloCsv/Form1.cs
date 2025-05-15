using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;

namespace ExceloCsv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            // Requerido por ExcelDataReader para registros regionales
            System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnCargarArchivo_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel o CSV|*.xlsx;*.csv";
            ofd.Title = "Selecciona un archivo";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string extension = Path.GetExtension(ofd.FileName).ToLower();

                if (extension == ".csv")
                {
                    CargarCSV(ofd.FileName);
                }
                else if (extension == ".xlsx")
                {
                    CargarExcel(ofd.FileName);
                }
                else
                {
                    MessageBox.Show("Formato de archivo no soportado.");
                }
            }
        }
        private void CargarCSV(string ruta)
        {
            listViewDatos.Items.Clear();
            listViewDatos.Columns.Clear();

            string[] lineas = File.ReadAllLines(ruta, Encoding.UTF8);

            if (lineas.Length == 0)
            {
                MessageBox.Show("El archivo está vacío.");
                return;
            }

            string[] encabezados = lineas[0].Split(',');
            foreach (var encabezado in encabezados)
            {
                listViewDatos.Columns.Add(encabezado.Trim());
            }

            for (int i = 1; i < lineas.Length; i++)
            {
                string[] celdas = lineas[i].Split(',');
                ListViewItem item = new ListViewItem(celdas[0]);

                for (int j = 1; j < celdas.Length; j++)
                {
                    item.SubItems.Add(celdas[j]);
                }

                listViewDatos.Items.Add(item);
            }

            listViewDatos.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void CargarExcel(string ruta)
        {
            listViewDatos.Items.Clear();
            listViewDatos.Columns.Clear();

            using (var stream = File.Open(ruta, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var resultado = reader.AsDataSet();

                // Usamos la primera hoja del Excel
                DataTable tabla = resultado.Tables[0];

                // Encabezados
                foreach (DataColumn col in tabla.Columns)
                {
                    listViewDatos.Columns.Add(col.ColumnName);
                }

                // Filas
                foreach (DataRow fila in tabla.Rows)
                {
                    ListViewItem item = new ListViewItem(fila[0]?.ToString());

                    for (int i = 1; i < tabla.Columns.Count; i++)
                    {
                        item.SubItems.Add(fila[i]?.ToString());
                    }

                    listViewDatos.Items.Add(item);
                }

                listViewDatos.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}