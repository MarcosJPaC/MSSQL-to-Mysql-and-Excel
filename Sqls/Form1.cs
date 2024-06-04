using System;
using System.Data;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using System.Collections.Generic;

namespace Sqls
{
    public partial class Form1 : Form
    {
        //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        private string sqlServerConnectionString = "Server=LAPTOP-MAIS200N\\SQLEXPRESS;Integrated Security=True;TrustServerCertificate=True;";
        private string mysqlConnectionString = "Server=localhost;User Id=root;Password=abc;";

        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            InitializeControls();
        }

        private void InitializeControls()
        {
            radioButton1.Enabled = false;
            radioButton2.Enabled = false;
            button1.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadDatabases();
        }

        private void LoadDatabases()
        {
            DataTable databasesTable;
            using (SqlConnection serverConnection = new SqlConnection(sqlServerConnectionString))
            {
                serverConnection.Open();
                databasesTable = serverConnection.GetSchema("Databases");
                serverConnection.Close();
            }

            foreach (DataRow row in databasesTable.Rows)
            {
                comboBox1.Items.Add(row["database_name"].ToString());
            }
        }

        private void comboBoxDatabases_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
        }

        private void ExportToExcel(string sqlServerDbConnectionString, string selectedDatabase)
        {
            string excelFilePath = $@"C:\Users\marco\OneDrive\Escritorio\converter\{selectedDatabase}.xlsx";

            if (File.Exists(excelFilePath))
            {
                DialogResult dialogResult = MessageBox.Show("El archivo ya existe. ¿Deseas sobrescribirlo?", "Archivo existente", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Guardar archivo como";
                    saveFileDialog.FileName = $"{selectedDatabase}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        excelFilePath = saveFileDialog.FileName;
                    }
                    else
                    {
                        return;
                    }
                }
            }

            using (SqlConnection sqlConnection = new SqlConnection(sqlServerDbConnectionString))
            {
                sqlConnection.Open();
                DataTable schemaTable = sqlConnection.GetSchema("Tables");

                using (ExcelPackage package = new ExcelPackage())
                {
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        string tableName = row["TABLE_NAME"].ToString();
                        string query = $"SELECT * FROM [{tableName}]"; // Escapar el nombre de la tabla

                        using (SqlCommand command = new SqlCommand(query, sqlConnection))
                        {
                            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                            {
                                DataTable dataTable = new DataTable();
                                adapter.Fill(dataTable);

                                string worksheetName = tableName;
                                int count = 1;
                                while (package.Workbook.Worksheets[worksheetName] != null)
                                {
                                    worksheetName = $"{tableName}_{count}";
                                    count++;
                                }

                                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
                                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                            }
                        }
                    }

                    FileInfo excelFile = new FileInfo(excelFilePath);
                    package.SaveAs(excelFile);
                }
            }

            MessageBox.Show("Datos exportados exitosamente a " + excelFilePath);
        }

        private void ExportToMySql(string sqlServerDbConnectionString, string selectedDatabase)
        {
            using (SqlConnection sqlConnection = new SqlConnection(sqlServerDbConnectionString))
            using (MySqlConnection mysqlConnection = new MySqlConnection(mysqlConnectionString))
            {
                sqlConnection.Open();
                mysqlConnection.Open();

                // Crear la base de datos en MySQL si no existe
                string createDatabaseQuery = $"CREATE DATABASE IF NOT EXISTS `{selectedDatabase}`;";
                using (MySqlCommand createDbCommand = new MySqlCommand(createDatabaseQuery, mysqlConnection))
                {
                    createDbCommand.ExecuteNonQuery();
                }

                // Seleccionar la base de datos recién creada
                mysqlConnection.ChangeDatabase(selectedDatabase);

                DataTable schemaTable = sqlConnection.GetSchema("Tables");

                foreach (DataRow row in schemaTable.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();
                    string query = $"SELECT * FROM [{tableName}]"; // Escapar el nombre de la tabla

                    using (SqlCommand command = new SqlCommand(query, sqlConnection))
                    {
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            using (MySqlCommand mySqlCommand = new MySqlCommand($"DROP TABLE IF EXISTS `{tableName}`;", mysqlConnection)) // Escapar el nombre de la tabla
                            {
                                mySqlCommand.ExecuteNonQuery();
                            }

                            string createTableQuery = GenerateCreateTableQuery(tableName, dataTable);
                            using (MySqlCommand mySqlCommand = new MySqlCommand(createTableQuery, mysqlConnection))
                            {
                                mySqlCommand.ExecuteNonQuery();
                            }

                            foreach (DataRow dataRow in dataTable.Rows)
                            {
                                string insertQuery = GenerateInsertQuery(tableName, dataRow);
                                using (MySqlCommand mySqlCommand = new MySqlCommand(insertQuery, mysqlConnection))
                                {
                                    // Asignar el valor de los datos binarios al parámetro
                                    foreach (DataColumn column in dataTable.Columns)
                                    {
                                        if (column.DataType == typeof(byte[]))
                                        {
                                            mySqlCommand.Parameters.AddWithValue("@BinaryData", dataRow[column]);
                                            break;
                                        }
                                    }

                                    mySqlCommand.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                }
            }

            MessageBox.Show("Datos exportados exitosamente a MySQL.");
        }

        private string GenerateCreateTableQuery(string tableName, DataTable dataTable)
        {
            string createTableQuery = $"CREATE TABLE `{tableName}` (";

            foreach (DataColumn column in dataTable.Columns)
            {
                string columnName = column.ColumnName;
                string dataType = GetMySQLDataType(column.DataType);

                createTableQuery += $"{columnName} {dataType}, ";
            }

            createTableQuery = createTableQuery.TrimEnd(',', ' ') + ");";
            return createTableQuery;
        }


        private string GetMySQLDataType(Type dataType)
        {
            if (dataType == typeof(int))
                return "INT";
            else if (dataType == typeof(string))
                return "VARCHAR(255)";
            else if (dataType == typeof(DateTime))
                return "DATETIME";
            else if (dataType == typeof(decimal))
                return "DECIMAL(18,2)";
            else if (dataType == typeof(bool))
                return "BIT";
            else if (dataType == typeof(short))
                return "SMALLINT";
            else if (dataType == typeof(float))
                return "FLOAT";
            else if (dataType == typeof(byte[])) // Cambiado a BLOB para MariaDB
                return "BLOB"; // Utilizamos BLOB para representar datos binarios en MariaDB
            throw new ArgumentException($"Tipo de datos no compatible: {dataType}");
        }



        private string EscapeSingleQuotes(string input)
        {
            return input.Replace("'", "''");
        }

        private string GenerateInsertQuery(string tableName, DataRow dataRow)
        {
            string insertQuery = $"INSERT INTO `{tableName}` VALUES (";

            List<string> values = new List<string>();

            foreach (var item in dataRow.ItemArray)
            {
                string valueToAdd;
                if (item is bool)
                {
                    // Convertir el valor booleano a 0 o 1 para MySQL
                    valueToAdd = ((bool)item) ? "1" : "0";
                }
                else if (item is decimal)
                {
                    valueToAdd = ((decimal)item).ToString(CultureInfo.InvariantCulture); // Convertir el decimal sin cambiar las comas
                }
                else if (item is string)
                {
                    string stringValue = (string)item;

                    // Limitar la longitud de la cadena (por ejemplo, 255 caracteres)
                    int maxLength = 255;
                    if (stringValue.Length > maxLength)
                    {
                        // Si la cadena es más larga que el máximo permitido, truncarla
                        stringValue = stringValue.Substring(0, maxLength);
                    }

                    // Escapar las comillas simples dentro del valor de la cadena
                    valueToAdd = EscapeSingleQuotes(stringValue);
                    valueToAdd = $"'{valueToAdd}'"; // Agregar comillas alrededor del valor escapado
                }
                else if (item == DBNull.Value) // Manejar valores nulos
                {
                    valueToAdd = "NULL";
                }
                else if (item is DateTime)
                {
                    valueToAdd = ((DateTime)item).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    valueToAdd = $"'{valueToAdd}'"; // Agregar comillas alrededor de la fecha formateada
                }
                else if (item is short) // Manejar el tipo de datos System.Int16
                {
                    // Convertir System.Int16 a un entero válido en MySQL
                    valueToAdd = ((short)item).ToString();
                }
                else if (item is byte[]) // Manejar el tipo de datos byte[]
                {
                    byte[] byteArray = (byte[])item;

                    // Limitar la cantidad de bytes a insertar (por ejemplo, 1000)
                    int maxLength = 1000;
                    if (byteArray.Length > maxLength)
                    {
                        // Si los datos binarios son más largos que el máximo permitido, truncarlos
                        byte[] truncatedArray = new byte[maxLength];
                        Array.Copy(byteArray, truncatedArray, maxLength);
                        byteArray = truncatedArray; // asignar el array truncado a byteArray
                    }

                    // Convertir los datos binarios en una cadena hexadecimal para la inserción
                    string hexString = BitConverter.ToString(byteArray).Replace("-", "");
                    valueToAdd = $"'{hexString}'"; // Agregar comillas alrededor de la cadena hexadecimal
                }
                else
                {
                    valueToAdd = item.ToString();
                }

                values.Add(valueToAdd);
            }

            insertQuery += string.Join(", ", values) + ");";
            return insertQuery;
        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            radioButton1.Enabled = true;
            radioButton2.Enabled = true;
            button1.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Por favor seleccione una base de datos.");
                return;
            }

            string selectedDatabase = comboBox1.SelectedItem.ToString();
            string sqlServerDbConnectionString = $"Server=LAPTOP-MAIS200N\\SQLEXPRESS;Database={selectedDatabase};Integrated Security=True;TrustServerCertificate=True;";

            if (radioButton1.Checked)
            {
                ExportToExcel(sqlServerDbConnectionString, selectedDatabase);
            }
            else if (radioButton2.Checked)
            {
                ExportToMySql(sqlServerDbConnectionString, selectedDatabase);
            }
        }
    }
}
