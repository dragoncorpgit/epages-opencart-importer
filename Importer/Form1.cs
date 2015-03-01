using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;
namespace Exportar
{

    public partial class Form1 : Form
    {
        public int maxManufacturer_ID, manufacturer_id, weight_class_id, tax_class_id, numberOfLinesInTheFile, numberTested = 0;
        private string price, weight;
        private string imageAlreadyInDatabase, databaseProductDescriptionEqual, product_id_AlreadyInDatabase, model_AlreadyInDatabase;
        private List<string> manufacturerExisting = new List<string>();
        private long maxImage_id, currentProduct_id, currentImage_id;
        private static Microsoft.Office.Interop.Excel.Application appExcel;
        private static Workbook newWorkbook = null;
        private static _Worksheet objsheet = null;
        private List<string> productsNameThatAreCorkList = new List<string>();
        private List<string> productNameTested = new List<string>();
        private const string folderToGetImages = old folder; //folder to get old images to import 
        private const string folderToPutImages = new_folder //folder to put images (server folder)
        private const int reduceTaxID = 11;
        private const int basicTaxID = 9;
        private const string host = your host;
        private const string port = your mysql server port;
        private const string dbUsername = your username;
        private const string dbName = your database name;
        private const string dbPassword = your password;
        public Form1()
        {
            InitializeComponent();
        }

        //Method to initialize opening Excel
        static void excel_init(String path)
        {
            appExcel = new Microsoft.Office.Interop.Excel.Application();

            if (System.IO.File.Exists(path))
            {
                // then go and load this into excel
                newWorkbook = appExcel.Workbooks.Open(path, true, true);
                objsheet = (_Worksheet)appExcel.ActiveWorkbook.ActiveSheet;
            }
            else
            {
                MessageBox.Show("Unable to open file!");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
                appExcel = null;
                System.Windows.Forms.Application.Exit();
            }

        }

        //Method to get value; cellname is A1,A2, or B1,B2 etc...in excel.
        static string excel_getValue(string cellname)
        {
            string value = string.Empty;
            try
            {
                value = objsheet.get_Range(cellname).get_Value().ToString();
            }
            catch
            {
                value = "";
            }

            return value;
        }

        //Method to close excel connection
        static void excel_close()
        {
            if (appExcel != null)
            {
                try
                {
                    newWorkbook.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
                    appExcel = null;
                    objsheet = null;
                }
                catch (Exception ex)
                {
                    appExcel = null;
                    MessageBox.Show("Unable to release the Object " + ex.ToString());
                }
                finally
                {
                    GC.Collect();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.IO.StreamReader sr;
            string productName, productName_PT, productDescription, productDescription_PT, productMetaDescription, productMetaDescription_PT, meta_keyword, meta_keyword_PT;
            int productAddToDatabase = 0, productEnabled = 1;
            string model;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    sr = new System.IO.StreamReader(openFileDialog1.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error - " + ex.Message);
                    return;
                }
                if (nupLinesImport.Value == 0)
                {
                    GetTheNumberOfLinesInTheFile(openFileDialog1, "N");
                }
                else
                {
                    numberOfLinesInTheFile = (int)nupLinesImport.Value;
                }
                excel_init(openFileDialog1.InitialDirectory + openFileDialog1.FileName);
                MessageBox.Show("File open");

                //  PopulateOptiniosOfProducts();
                try
                {
                    PopulateManufacturer();
                    MySqlConnection con = ServerConnect();
                    con.Open();
                    for (int i = 2; i < numberOfLinesInTheFile; i++)
                    {
                        model = excel_getValue("N" + i);


                        MySql.Data.MySqlClient.MySqlCommand myCommand;
                        //MySqlDataReader reader = myCommand.ExecuteReader();
                        MySqlDataReader reader;
                        int max_Product_id = 0;
                        /*
                            while (reader.Read())
                            {
                                max_Product_id = Int32.Parse(reader[0].ToString());

                            }
                            reader.Close();
    */
                        int isVisible, weight_Class, quantity, stock_status_id, shipping, points, length, width, height, viewed, lenght_class_id;

                        string imageUrl, date_available, date_added, date_modified;



                        string tempStringIsvisible = excel_getValue("P" + i);
                        isVisible = Int16.Parse(tempStringIsvisible);


                        if (excel_getValue("AK" + i) == "gram")
                        {
                            weight_Class = 2;

                        }
                        else
                        {
                            weight_Class = 1;
                        }
                        string tempStringWeight = excel_getValue("AL" + i);
                        if (tempStringWeight != "")
                        {
                            weight = tempStringWeight;
                        }


                        currentProduct_id = max_Product_id + 1;

                        quantity = 1;
                        stock_status_id = 7;
                        string tempStrigImage = excel_getValue("EH" + i);
                        if (tempStrigImage != "")
                        {
                            string imageName = tempStrigImage;
                            imageUrl = "data/products/" + imageName;
                            if (!GetimageOfProduct(imageName))
                            {
                                imageUrl = "";
                                productEnabled = 0;
                            }
                        }
                        else
                        {
                            imageUrl = "";
                            productEnabled = 0;
                        }


                        string tempStringManufacter = excel_getValue("AE" + i);
                        if (tempStringManufacter != "")
                        {
                            MySql.Data.MySqlClient.MySqlCommand manufacturerQuery = new MySql.Data.MySqlClient.MySqlCommand
                                ("SELECT manufacturer_id FROM oc_manufacturer  WHERE name = '" + tempStringManufacter + "'", con);
                            MySqlDataReader manufacturerReader = manufacturerQuery.ExecuteReader();

                            while (manufacturerReader.Read())
                            {
                                manufacturer_id = Int32.Parse(manufacturerReader[0].ToString());
                            }

                            manufacturerReader.Close();
                        }
                        else
                        {
                            manufacturer_id = 0;
                        }

                        shipping = 1;
                        string priceString = excel_getValue("Q" + i);
                        if (priceString != "")
                        {
                            if (priceString != "0")
                            {
                                price = priceString.Replace(',', '.');
                            }
                            else
                            {
                                price = "0";
                            }
                        }

                        string excelProductTax = excel_getValue("W" + i);

                        if (excelProductTax == "normal")
                        {
                            tax_class_id = basicTaxID;
                        }
                        else if (excelProductTax == "reduced")
                        {
                            tax_class_id = reduceTaxID;
                        }
                        else
                        {
                            MessageBox.Show("Problem with product taxes");
                        }

                        date_available = System.DateTime.Now.ToString("yyyy.MM.dd");
                        length = 0;
                        height = 0;
                        width = 0;
                        lenght_class_id = 1;
                        date_added = date_available;
                        date_modified = date_available;
                        viewed = 0;

                        /*     myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT model FROM oc_product WHERE model ='" + model + "'", con);
                             reader = myCommand.ExecuteReader();
                             model_AlreadyInDatabase = "";
                             while (reader.Read())
                             {
                                 model_AlreadyInDatabase = reader[0].ToString();
                             }
                             reader.Close();
                             if (model_AlreadyInDatabase == "")
                             {
                            */
                        productAddToDatabase++;
                        myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_product (model, quantity, stock_status_id, image, manufacturer_id, shipping, price, tax_class_id," +
                    "date_available, weight, weight_class_id, length, width, height, length_class_id, subtract, minimum, sort_order, status, viewed, date_added, date_modified) VALUES ('" + model
                            + "','" + quantity + "', '" + stock_status_id + "', '" + imageUrl + "', '" + manufacturer_id + "', '" + shipping + "', '" + price + "', '" + tax_class_id
                            + "', '" + date_available + "', '" + weight + "','" + weight_class_id + "', '" + length + "', '" + width + "', '" + height + "', '"
                            + lenght_class_id + "', '0', '1', '0', '" + productEnabled + "', '" + viewed + "','" + date_added + "', '" + date_modified + "')", con);
                        myCommand.ExecuteNonQuery();
                        currentProduct_id = myCommand.LastInsertedId;
                        //}

                        //finish oc_products


                        string listOfImages = excel_getValue("DW" + i);

                        string[] imagesInDetailName = listOfImages.Split(';');

                        for (int i2 = 0; i2 < imagesInDetailName.Length; i2++)
                        {


                            if (imagesInDetailName[i2] != "")
                            {
                                if (!GetimageOfProduct(imagesInDetailName[i2]))
                                {
                                    continue;
                                }

                                imagesInDetailName[i2] = "data/products/" + imagesInDetailName[i2];
                                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT MAX(product_image_id) FROM oc_product_image", con);

                                reader = myCommand.ExecuteReader();
                                string tempImage_ID = "";
                                while (reader.Read())
                                {
                                    tempImage_ID = reader[0].ToString();
                                }
                                if (tempImage_ID == "")
                                {
                                    maxImage_id = 0;
                                }
                                else
                                {
                                    maxImage_id = int.Parse(tempImage_ID);
                                }

                                reader.Close();
                                currentImage_id = maxImage_id + 1;

                                /*    myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT product_id FROM oc_product_image WHERE image='" + imagesInDetailName[i2]
                                        + "' AND product_id ='" + currentProduct_id + "' ", con);

                                    reader = myCommand.ExecuteReader();
                                    imageAlreadyInDatabase = "";
                                    while (reader.Read())
                                    {
                                        imageAlreadyInDatabase = reader[0].ToString();
                                    }
                                    reader.Close();
                                   

                                    if (imageAlreadyInDatabase == "")
                                    {
                                 */
                                myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_product_image VALUES ('" + currentImage_id + "','"
                                    + currentProduct_id + "','" + imagesInDetailName[i2] + "', '') ", con);
                                myCommand.ExecuteNonQuery();

                                //  }
                            }

                        }

                        //finish oc_product_image

                        productName = excel_getValue("AQ" + i);
                        productName_PT = excel_getValue("AZ" + i);
                        productDescription = excel_getValue("CE" + i);
                        productDescription_PT = excel_getValue("CN" + i);
                        productMetaDescription = productName;
                        productMetaDescription_PT = productName_PT;
                        meta_keyword = excel_getValue("CO" + i);
                        meta_keyword_PT = excel_getValue("CX" + i);
                        if (productName == "")
                        {
                            productName = "Missing name ";

                        }

                        if (productDescription == "")
                        {
                            productDescription = "Description missing";
                        }

                        productName = productName.Replace("'", "");
                        productDescription = productDescription.Replace("'", "");
                        productMetaDescription = productMetaDescription.Replace("'", "");

                        /*

                        myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT product_id FROM oc_product_description WHERE description ='"
                            + productDescription + "' AND name='" + productName + "'", con);
                        reader = myCommand.ExecuteReader();
                        databaseProductDescriptionEqual = "";
                        while (reader.Read())
                        {
                            databaseProductDescriptionEqual = reader[0].ToString();
                        }
                        reader.Close();
                        if (databaseProductDescriptionEqual == "")
                        {
                         * */

                        productName = productName.Replace("'", "''");
                        productDescription = productDescription.Replace("'", "''");
                        productMetaDescription = productMetaDescription.Replace("'", "''");
                        meta_keyword = meta_keyword.Replace("'", "''");

                        /*   productName = productName.Replace("'s", "''s");
                           productDescription = productDescription.Replace("'s", "''s");
                           productMetaDescription = productMetaDescription.Replace("'s", "''s");
                           meta_keyword = meta_keyword.Replace("'s", "''s");
                             */
                        myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_product_description VALUES ('" + currentProduct_id + "', '1', '"
                            + productName + "', '" + productDescription + "', '', '" + productMetaDescription + "', '', '" + meta_keyword + "')", con);
                        myCommand.ExecuteNonQuery();

                        productName_PT = productName_PT.Replace("'", "''");
                        productDescription_PT = productDescription_PT.Replace("'", "''");
                        productMetaDescription_PT = productMetaDescription_PT.Replace("'", "''");
                        meta_keyword_PT = meta_keyword_PT.Replace("'", "''");


                        myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_product_description VALUES ('" + currentProduct_id + "', '2', '"
                               + productName_PT + "', '" + productDescription_PT + "', '', '" + productMetaDescription_PT + "', '', '" + meta_keyword_PT + "')", con);
                        myCommand.ExecuteNonQuery();


                        //    }

                        //finish oc_product_description

                        /*   myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT product_id FROM  oc_product_to_store WHERE product_id='"
                               + currentProduct_id + "' ", con);
                           reader = myCommand.ExecuteReader();


                           product_id_AlreadyInDatabase = "";
                           while (reader.Read())
                           {
                               product_id_AlreadyInDatabase = reader[0].ToString();
                           }
                           reader.Close();

                           if (product_id_AlreadyInDatabase == "")
                           {
                         */
                        myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_product_to_store VALUES ( '" + currentProduct_id + "','0') ", con);
                        myCommand.ExecuteNonQuery();
                        //    }
                        //finish oc_product_to_store

                        maxManufacturer_ID = 0; manufacturer_id = 0; weight_class_id = 0; tax_class_id = 0; weight = "0";
                        price = "0";
                        imageAlreadyInDatabase = ""; productName = ""; productDescription = ""; databaseProductDescriptionEqual = ""; product_id_AlreadyInDatabase = ""; model_AlreadyInDatabase = "";

                        maxImage_id = 0; currentProduct_id = 0; currentImage_id = 0;
                        productEnabled = 1;

                    }
                    con.Close();
                    MessageBox.Show("Finish - Products Add -" + productAddToDatabase);

                    sr.Close();

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private MySqlConnection ServerConnect()
        {
            MySqlConnection con = new MySqlConnection("Server=" + host +"; Port=" + port +"; Database=" + dbName +";UID=" + dbUsername +";Password=" + dbPassword);
            return con;
        }
        private void PopulateManufacturer()
        {

            MySqlConnection con = ServerConnect();
            MySql.Data.MySqlClient.MySqlCommand myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT MAX( manufacturer_id) FROM oc_manufacturer", con);
            if (con.State != ConnectionState.Open)
            {
                con.Open();
                MySqlDataReader reader = myCommand.ExecuteReader();


                while (reader.Read())
                {
                    maxManufacturer_ID = Int32.Parse(reader[0].ToString());
                }
                reader.Close();
                con.Close();

                con.Open();
                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT name FROM oc_manufacturer", con);
                reader = myCommand.ExecuteReader();
              
                while (reader.Read())
                {
                    manufacturerExisting.Add(reader[0].ToString());
                }
                reader.Close();
                con.Close();

                for (int i = 2; i < numberOfLinesInTheFile; i++)
                {

                    string manufacturer = excel_getValue("AE" + i);
                    if (manufacturer != "" )
                    {

                        if (manufacturerExisting.IndexOf(manufacturer) == -1)
                        {

                            con.Open();

                            myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_manufacturer VALUES ('" + (maxManufacturer_ID + i) + "','" + manufacturer + "', '', '0' )", con);
                            myCommand.ExecuteNonQuery();

                            con.Close();
                        }
                    }
                }

            }
           
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string categoryName = "", categoryNamePT = "", categoryAlreadyInDatabase = "", parentName = "", metaTitle = "", metaTitlePT, metaKeyword = "", metaKeywordPT = "";
            long currentCategory_ID = 0;
            int parent_ID = 0, categoryAddToDatabase = 0;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                System.IO.StreamReader sr = new
                System.IO.StreamReader(openFileDialog1.FileName);
                GetTheNumberOfLinesInTheFile(openFileDialog1, "C");
                excel_init(openFileDialog1.InitialDirectory + openFileDialog1.FileName);

                MySqlConnection con = ServerConnect();
                string date_addedAndModified = System.DateTime.Now.ToString("yyyy.MM.dd");
               

                try
                {
                    con.Open();
                    for (int i = 2; i < numberOfLinesInTheFile; i++)
                    {
                        if (excel_getValue("A" + i) == "Category")
                        {

                            string tempCategoryFullName = excel_getValue("B" + i).Replace("/Malas_em_Cortica", ""); //cork fix
                            string[] categoryPath = tempCategoryFullName.Split('/');

                            bool isParent = false;
                            int maxCategory_ID = 0, top = 0;

                            categoryName = excel_getValue("G" + i);

                            if (categoryName == "")
                            {
                                categoryName = excel_getValue("C" + i);

                            }

                            categoryNamePT = excel_getValue("P" + i);
                            metaTitle = categoryName;
                            metaTitlePT = categoryNamePT;
                            metaKeyword = excel_getValue("BO" + i);
                            metaKeyword = metaKeyword.Replace(" ", ", ");
                            metaKeywordPT = excel_getValue("BX" + i);
                            try
                            {
                                parentName = categoryPath[(categoryPath.Length - 2)].ToString();
                            }
                            catch
                            {
                                parentName = categoryPath[(categoryPath.Length - 1)].ToString();
                            }
                            parentName = parentName.Replace("\"", " ");
                          
                            MySqlDataReader reader = null;
                            MySql.Data.MySqlClient.MySqlCommand myCommand;


                            if (categoryName.Contains("\""))
                            {
                                MessageBox.Show("ERROR");
                                return;
                            }
                         
                            myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT category_id FROM oc_category_description WHERE name='"
                                + parentName + "'", con);
                            reader = myCommand.ExecuteReader();
                            string stringRead = "";
                            while (reader.Read())
                            {
                                stringRead = reader[0].ToString();

                            }
                            reader.Close();
                            if (stringRead != "")
                            {
                                parent_ID = int.Parse(stringRead);
                                top = 1;
                            }
                            else
                            {
                                top = 0;
                                parent_ID = 0;
                            }

                            MySqlCommand comm = con.CreateCommand();
                            comm.CommandText = "SELECT category_id FROM oc_category_description WHERE name=?categoryName";
                            comm.Parameters.Add("?categoryName", categoryName);
                          
                            reader = comm.ExecuteReader();

                            while (reader.Read())
                            {
                                categoryAlreadyInDatabase = reader[0].ToString();
                            }
                            reader.Close();
                            

                                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT category_id FROM oc_category WHERE category_id='" + categoryAlreadyInDatabase +
                                    "'AND parent_id ='" + parent_ID + "'", con);
                                reader = myCommand.ExecuteReader();


                                while (reader.Read())
                                {
                                    categoryAlreadyInDatabase = reader[0].ToString();
                                }
                                reader.Close();

                                if (categoryAlreadyInDatabase != "")
                                {
                                    
                                    continue;
                                }
                            
                                    myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_category (`image`, `parent_id`, `top`, `column`, `sort_order`, `status`, `date_added`, `date_modified`) VALUES ('','"
                                        + parent_ID + "','" + top + "','0','0','1','" + date_addedAndModified + "','" + date_addedAndModified + "')", con);
                                    myCommand.ExecuteNonQuery();
                                    currentCategory_ID = myCommand.LastInsertedId;

                                  
                                    myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_category_description (category_id, language_id, name, meta_title, meta_keyword) VALUES ('" + currentCategory_ID + "', '1', '"
                                        + categoryName + "', '" + metaTitle +"', '" + metaKeyword +"')", con);
                                    myCommand.ExecuteNonQuery();

                                    myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_category_description (category_id, language_id, name, meta_title, meta_keyword) VALUES ('" + currentCategory_ID + "', '2', '"
                                        + categoryNamePT + "', '" + metaKeywordPT +"', '" + metaKeywordPT +"')", con);
                                    myCommand.ExecuteNonQuery();

                                    myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_category_to_store VALUES ('" + currentCategory_ID + "','0')", con);
                                    myCommand.ExecuteNonQuery();

                                    myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_category_path VALUES ('" + currentCategory_ID + "','" + currentCategory_ID + "', '0')", con);
                                    myCommand.ExecuteNonQuery();
                                    categoryAddToDatabase++;
                                
                            reader.Close();
                          
                           
                        }

                         categoryName = "";
                        categoryNamePT = ""; 
                        categoryAlreadyInDatabase = ""; 
                            parentName = ""; metaTitle = ""; 
                        metaTitlePT = ""; 
                        metaKeyword = ""; metaKeywordPT = "";
                         currentCategory_ID = 0;
                         parent_ID = 0;
                        categoryAddToDatabase = 0;
                       
                    }

                    con.Close();
                    MessageBox.Show("Has been export - " + categoryAddToDatabase);

                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    sr.Close();
                    excel_close();
                }
               
            }
        }

        private bool hasRelateProductAndCategory(string currentProductRelated, string currentCategoryName, MySqlConnection con)
        {

            MySqlDataReader reader = null;
            MySql.Data.MySqlClient.MySqlCommand myCommand;
            int product_IDToGiveACategory = 0;
            int category_ID = 0;
            try
            {
                if (!(con.State == ConnectionState.Open))
                {
                    con.Open();
                }


                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT product_id FROM oc_product WHERE model='" + currentProductRelated + "'", con);
                reader = myCommand.ExecuteReader();

                while (reader.Read())
                {
                    int.TryParse(reader[0].ToString(), out product_IDToGiveACategory);
                }

                reader.Close();
                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT category_id FROM oc_category_description WHERE name='" + currentCategoryName + "'", con);
                reader = myCommand.ExecuteReader();
                while (reader.Read())
                {
                    int.TryParse(reader[0].ToString(), out category_ID);

                }
                reader.Close();
                if (category_ID == 0 || product_IDToGiveACategory == 0)
                {
                    return false;
                }

                string relationAlreadyExist = "";
                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT product_id FROM oc_product_to_category WHERE product_id ='" + product_IDToGiveACategory +
                    "' AND category_id = '" + category_ID + "'", con);
                reader = myCommand.ExecuteReader();

                while (reader.Read())
                {
                    relationAlreadyExist = reader[0].ToString();
                }
                reader.Close();
                if (relationAlreadyExist == "")
                {
                    myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_product_to_category VALUES ('" + product_IDToGiveACategory + "','"
                        + category_ID + "')", con);
                    myCommand.ExecuteNonQuery();
                }

               
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
            finally
            {
                con.Close();
               
            }
           

        }
        private void GetTheNumberOfLinesInTheFile(OpenFileDialog OpenFileDialog1, string ColumnThatIsNotEmpty)
        {
            excel_init(openFileDialog1.InitialDirectory + openFileDialog1.FileName);

            for (int i = 1; i < 99999; i++)
            {

                if (excel_getValue(ColumnThatIsNotEmpty + i) == "")
                {

                    numberOfLinesInTheFile = i;
                    i = 99999;

                }
            }
            excel_close();
        }

        private void PopulateOptiniosOfProducts()
        {

            MySqlConnection con = ServerConnect();
            con.Open();
            MySqlDataReader reader = null;
            MySql.Data.MySqlClient.MySqlCommand myCommand;

            for (int i = 1; i < numberOfLinesInTheFile; i++)
            {

                string nameOfVariation = excel_getValue("DM" + i), nameOfOption = "", currentColumn = "DM", maxColumn = textBox1.Text;
                int option_ID = 0;
                int[] valuerInAsciiOfColumns = new int[2];
                int[] maxValuerInAsciiOfColumns = new int[2];
                string currentValuerOfColumInString = "", subOption = "";
                int i1 = 0;
                foreach (char c in maxColumn)
                {

                    maxValuerInAsciiOfColumns[i1] = System.Convert.ToInt32(c);
                    i1++;
                }

                if ((nameOfVariation == "cor") || (nameOfVariation == "Cor") || (nameOfVariation == "Cores"))
                {

                    nameOfOption = "Color";
                }


                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT option_id FROM oc_option_description WHERE name ='" + nameOfOption + "'", con);

                reader = myCommand.ExecuteReader();

                while (reader.Read())
                {
                    option_ID = int.Parse(reader[0].ToString());
                }
                reader.Close();

                int i2 = 0;
                foreach (char c in currentColumn)
                {

                    valuerInAsciiOfColumns[i2] = System.Convert.ToInt32(c);
                    i2++;
                }

                while (currentValuerOfColumInString != maxColumn)
                {


                    if (valuerInAsciiOfColumns[1] < 90)
                    {
                        valuerInAsciiOfColumns[1]++;
                    }
                    else
                    {
                        valuerInAsciiOfColumns[0]++;
                        valuerInAsciiOfColumns[0] = 65;
                    }


                    currentValuerOfColumInString = char.ConvertFromUtf32(valuerInAsciiOfColumns[0]) + char.ConvertFromUtf32(valuerInAsciiOfColumns[1]);


                    subOption = excel_getValue(currentValuerOfColumInString + i);

                    if (subOption != "")
                    {
                        currentValuerOfColumInString = maxColumn;

                    }
                }

                int currentOptionValue_ID = 0;

                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT MAX(option_value_id) FROM oc_option_value", con);
                reader = myCommand.ExecuteReader();

                while (reader.Read())
                {

                    currentOptionValue_ID = int.Parse(reader[0].ToString());
                }
                reader.Close();
                currentOptionValue_ID++;

                myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_option_value VALUES ('" + currentOptionValue_ID + "','"
                      + option_ID + "', '', '0')", con);
                myCommand.ExecuteNonQuery();

                string subOptionAlreadyInDatabase = "";

                myCommand = new MySql.Data.MySqlClient.MySqlCommand("SELECT name FROM oc_option_value_description WHERE name ='" + subOption + "'", con);
                reader = myCommand.ExecuteReader();

                while (reader.Read())
                {

                    subOptionAlreadyInDatabase = reader[0].ToString();
                }
                reader.Close();

                if (subOptionAlreadyInDatabase == "")
                {

                    myCommand = new MySql.Data.MySqlClient.MySqlCommand("INSERT INTO oc_option_value_description VALUES ('" + currentOptionValue_ID + "','1', '" + option_ID + "', '" + subOption + "')", con);
                    myCommand.ExecuteNonQuery();
                }

            }
          

        }

        private bool IsCork(string productNameToTest)
        {
            numberTested++;

            for (int i = 0; i < productsNameThatAreCorkList.Count; i++)
            {
                if (productNameToTest == productsNameThatAreCorkList[i])
                {
                    productNameTested.Add(productNameToTest);
                    productsNameThatAreCorkList.Remove(productNameToTest);
                    return true;

                }
            }

            return false;
        }


        private bool GetimageOfProduct(string imageName)
        {
            try
            {
                if (File.Exists(folderToPutImages + imageName))
                {
                    File.Delete(folderToPutImages + imageName);
                }
                File.Copy(folderToGetImages + imageName, folderToPutImages + imageName, false);
                return true;
            }
            catch (System.IO.IOException e)
            {
                return false;
            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            int productsRelated = 0;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                System.IO.StreamReader sr = new
                System.IO.StreamReader(openFileDialog1.FileName);
                
               
                if (nupLinesImport.Value == 0)
                {
                    GetTheNumberOfLinesInTheFile(openFileDialog1, "A");
                }
                else
                {
                    numberOfLinesInTheFile = (int) nupLinesImport.Value;
                }
                excel_init(openFileDialog1.InitialDirectory + openFileDialog1.FileName);
                 MySqlConnection con = ServerConnect();
                con.Open();
                for (int i = 2; i < numberOfLinesInTheFile; i++)
                {
                    string [] tempString = excel_getValue("A" + i).Split('/');
                    string categoryName = tempString[tempString.Length - 1];
                    categoryName = categoryName.Replace("Categories", "");
                    if (categoryName == "")
                    {
                        continue;
                    }
                    string productName = excel_getValue("B" + i);
                    productName = productName.Replace('"', '\0');
                    if (hasRelateProductAndCategory(productName, categoryName, con))
                    {
                        productsRelated++;
                    }
                   
                }
                con.Close();
                MessageBox.Show("Has been related " + productsRelated + " products");
                
            }
        }

    }
}
