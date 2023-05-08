using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Drawing;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Security.Cryptography;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style.XmlAccess;

namespace Projekt
{
    class PopulationStatistics
    {
        //Private fields (variables) of the class, used in the ReadFile function
        List<string[]> DataBase = new List<string[]>();  //List of population data from the file
        string[] columnNames = new string[15];           //Array of column names for the data in the DataBase list

        //Private fields (variables) of the class for calculations on data from the file
        string[] continentName = new string[6] { "Asia", "Africa", "SouthAmerica", "NorthAmerica", "Europe", "Oceania" };
        double[,] continentPopulation1970To2022 = new double[6, 8];
        double[] changeInPopulation1970To2022 = new double[6];
        double[] changeInPopulation2020To2022 = new double[6];
        double[] percentPopulationShareWorld = new double[6];
        double worldPopulation2022;

        //File path names
        string fileNameToRead;       //Path to the file to be read
        string fileNameExcel;        //Path name for creating an Excel file
        string fileNameToSave;       //Path to the file to be saved
        
        // Default constructor
        public PopulationStatistics()
        {
            fileNameToRead = "World Population.txt";
            fileNameExcel = "ProjectCharts.xlsx";
            fileNameToSave="Results.txt";
        }

        // Constructor that takes the file paths as an argument
        public PopulationStatistics(string file_name, string file_name_excel, string file_name_toSave)
        {
            fileNameToRead = file_name;
            fileNameExcel = file_name_excel;
            fileNameToSave = file_name_toSave;
        }
        // Function that reads data from the file into the DataBase list
        public void ReadFile()
        {
            try
            {
                if (Path.GetExtension(fileNameToRead) != ".txt")
                {
                    throw new ArgumentException("Invalid file type. The file type must be .txt.");
                }

                StreamReader reader = new StreamReader(fileNameToRead);  //Open the file for reading
                columnNames = reader.ReadLine().Split(';');  //Get column names from the first row of the file
                while (reader.EndOfStream == false)  //Read data and save it to the DataBase list
                {
                    string[] data = reader.ReadLine().Split(';');
                    for (int i = 5; i < 16; i++)
                    {
                        data[i] = data[i].Replace('.', ',');  //Modify string data for future conversion to double type
                    }
                    DataBase.Add(data);
                }
                reader.Close(); //Close the file
            }
            catch (ArgumentException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (FileNotFoundException) { Console.WriteLine("\nError: The specified file to read does not exist!"); }
            catch (IOException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (IndexOutOfRangeException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (Exception ex) { Console.WriteLine("\nError : " + ex.Message); }
        }
        // Function that reads data from the file into the continentPopulation1970To2020 aaray 
        public void ContinentPopulationFrom1970to2022()
        {
            try
            {
                for (int i = 0; i < DataBase.Count; i++)
                {
                    for (int j = 0; j < continentName.Length; j++)
                    {
                        if (DataBase[i][4] == continentName[j])
                        {
                            continentPopulation1970To2022[j, 0] += Convert.ToDouble(DataBase[i][12]);
                            continentPopulation1970To2022[j, 1] += Convert.ToDouble(DataBase[i][11]);
                            continentPopulation1970To2022[j, 2] += Convert.ToDouble(DataBase[i][10]);
                            continentPopulation1970To2022[j, 3] += Convert.ToDouble(DataBase[i][9]);
                            continentPopulation1970To2022[j, 4] += Convert.ToDouble(DataBase[i][8]);
                            continentPopulation1970To2022[j, 5] += Convert.ToDouble(DataBase[i][7]);
                            continentPopulation1970To2022[j, 6] += Convert.ToDouble(DataBase[i][6]);
                            continentPopulation1970To2022[j, 7] += Convert.ToDouble(DataBase[i][5]);
                            break;
                        }
                    }
                }
            }
            catch (FormatException) { Console.WriteLine("\nError : The format of the string does not correspond, the number of populations cannot be converted to type double! "); }
            catch (OverflowException) { Console.WriteLine("\nError : The number value is too large or too small for conversion to double."); }
            catch (IndexOutOfRangeException) { Console.WriteLine("\nError : The index to an array element is out of range!"); }
            catch (Exception ex) { Console.WriteLine("\nError : " + ex.Message); }
        }
        // Function that calculates the percentage change in population from 1970 to 2022 for each continent
        public void PercentChangePopulation1970To2022()
        {
            try
            {
                double[] continentPopulation1970 = new double[6];
                double[] continentPopulation2022 = new double[6];
                for (int i = 0; i < continentName.Length; i++)
                {
                    continentPopulation1970[i] += continentPopulation1970To2022[i, 0];
                    continentPopulation2022[i] += continentPopulation1970To2022[i, 7];
                }
                // Calculate percentage change in population 1970/2022
                for (int i = 0; i < continentPopulation1970.Length; i++)
                {
                    changeInPopulation1970To2022[i] = ((continentPopulation2022[i] * 100.0) / continentPopulation1970[i]) - 100.0;
                }
            }
            catch (IndexOutOfRangeException) { Console.WriteLine("\nError : The index to an array element is out of range!"); }
            catch (DivideByZeroException) { Console.WriteLine("\nError : Trying to divide by zero!"); }
            catch (Exception ex) { Console.WriteLine("\nError : " + ex.Message); }
        }
        // Function that calculates the percentage change in population from 2020 to 2022 for each continent
        public void PercentChangePopulation2020To2022()
        {
            try
            {
                double[] continentPopulation2020 = new double[6];
                double[] continentPopulation2022 = new double[6];
                for (int i = 0; i < continentName.Length; i++)
                {
                    continentPopulation2020[i] += continentPopulation1970To2022[i, 6];
                    continentPopulation2022[i] += continentPopulation1970To2022[i, 7];
                }
                // Calculate percentage change in population 2020/2022
                for (int i = 0; i < continentPopulation2020.Length; i++)
                {
                    changeInPopulation2020To2022[i] = ((continentPopulation2022[i] * 100.0) / continentPopulation2020[i]) - 100.0;
                }
            }
            catch (IndexOutOfRangeException) { Console.WriteLine("\nError : The index to an array element is out of range!"); }
            catch (DivideByZeroException) { Console.WriteLine("\nError : Trying to divide by zero!"); }
            catch (Exception ex) { Console.WriteLine("\nError : " + ex.Message); }
        }
        //Function that calculate percentage share of the population of each continent in the total world population in 2022
        public void PercentWorldPopulationShare2022()
        {
            try
            {
                // Calculate world population
                for (int i = 0; i < continentName.Length; i++)
                {
                    worldPopulation2022 += continentPopulation1970To2022[i, 7];
                }
                // Calculate percentage share of each continent's population in the world population
                for (int i = 0; i < continentName.Length; i++)
                {
                    percentPopulationShareWorld[i] = ((continentPopulation1970To2022[i, 7] * 100.0) / worldPopulation2022);
                }
            }
            catch (IndexOutOfRangeException) { Console.WriteLine("\nError : The index to an array element is out of range!"); }
            catch (DivideByZeroException) { Console.WriteLine("\nError : Trying to divide by zero!"); }
            catch (Exception ex) { Console.WriteLine("\nError : " + ex.Message); }
        }
        public void CreateChartsContinentPopulation()
        {
            try
            {
                if (Path.GetExtension(fileNameExcel) != ".xlsx")
                {
                    throw new ArgumentException("Invalid file type. The file type must be .xlsx.");
                }

                //create a new Excel file
                FileInfo newFile = new FileInfo(fileNameExcel);
                //create package
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage(newFile);
                //create worksheets
                ExcelWorksheet change_70_22 = package.Workbook.Worksheets.Add("Change_pop_1970_2022");
                ExcelWorksheet change_20_22 = package.Workbook.Worksheets.Add("Change_pop_2020_2022");
                ExcelWorksheet percent_world = package.Workbook.Worksheets.Add("Percentage_Share_World");
                ExcelWorksheet annual_change_70_22 = package.Workbook.Worksheets.Add("Annual_change_pop_1970_2022");

                //create data arrays in Excel
                change_70_22.Cells[1, 1].Value = "Continents";
                change_70_22.Cells[1, 2].Value = "Population_1970";
                change_70_22.Cells[1, 3].Value = "Population_2022";
                change_70_22.Cells[1, 4].Value = "Population_Change (%)";
                ////////////////////////////////////////////////////////////////////
                change_20_22.Cells[1, 1].Value = "Continents";
                change_20_22.Cells[1, 2].Value = "Population_2020";
                change_20_22.Cells[1, 3].Value = "Population_2022";
                change_20_22.Cells[1, 4].Value = "Population_Change (%)";
                ///////////////////////////////////////////////////////////////////
                percent_world.Cells[1, 1].Value = "Continents";
                percent_world.Cells[1, 2].Value = "Population_2022";
                percent_world.Cells[1, 3].Value = "Population_Percent_Share_World (%)";
                //////////////////////////////////////////////////////////////////
                annual_change_70_22.Cells[1, 1].Value = "Continents";
                annual_change_70_22.Cells[1, 2].Value = "1970";
                annual_change_70_22.Cells[1, 3].Value = "1980";
                annual_change_70_22.Cells[1, 4].Value = "1990";
                annual_change_70_22.Cells[1, 5].Value = "2000";
                annual_change_70_22.Cells[1, 6].Value = "2010";
                annual_change_70_22.Cells[1, 7].Value = "2015";
                annual_change_70_22.Cells[1, 8].Value = "2020";
                annual_change_70_22.Cells[1, 9].Value = "2022";

                //add data to Excel arrays
                for (int i = 0; i < continentName.Length; i++)
                {
                    change_70_22.Cells[i + 2, 1].Value = continentName[i];                     //Name of continent
                    change_70_22.Cells[i + 2, 2].Value = continentPopulation1970To2022[i, 0];  //Population by continent 1970
                    change_70_22.Cells[i + 2, 3].Value = continentPopulation1970To2022[i, 7];  //Population by continent 2022
                    change_70_22.Cells[i + 2, 4].Value = changeInPopulation1970To2022[i];
                }
                //////////////////////////////////////////////////////////////////////////////
                for (int i = 0; i < continentName.Length; i++)
                {
                    change_20_22.Cells[i + 2, 1].Value = continentName[i];                     //Name of continent
                    change_20_22.Cells[i + 2, 2].Value = continentPopulation1970To2022[i, 6];  //Population by continent 2020
                    change_20_22.Cells[i + 2, 3].Value = continentPopulation1970To2022[i, 7];  //Population by continent 2022
                    change_20_22.Cells[i + 2, 4].Value = changeInPopulation2020To2022[i];
                }
                /////////////////////////////////////////////////////////////////////////////
                for (int i = 0; i < continentName.Length; i++)
                {
                    percent_world.Cells[i + 2, 1].Value = continentName[i];                     //Name of continent
                    percent_world.Cells[i + 2, 2].Value = continentPopulation1970To2022[i, 6];  //Population by continent 2022
                    percent_world.Cells[i + 2, 3].Value = percentPopulationShareWorld[i];       //Poppulation share world
                }
                /////////////////////////////////////////////////////////////////////////////
                for (int i = 0; i < continentName.Length; i++)
                {
                    annual_change_70_22.Cells[i + 2, 1].Value = continentName[i];  //Name of continent
                    for (int j = 0; j < 8; j++)
                    {
                        annual_change_70_22.Cells[i + 2, j + 2].Value = continentPopulation1970To2022[i, j];  //Population by continent 1970-2022
                    }
                }

                // Adding charts based on data
                var graph_change_70_22 = change_70_22.Drawings.AddChart("Population change 1970-2022", eChartType.ColumnClustered3D);
                graph_change_70_22.SetPosition(0, 0, 10, 0); // (row,rowOffsetPixels,column,columnOffsetPixels) position of chart
                graph_change_70_22.SetSize(800, 600);        //chart size (Width,Hieght)

                var graph_per_change_70_22 = change_70_22.Drawings.AddChart("Population change (%) 1970-2022", eChartType.ColumnClustered);
                graph_per_change_70_22.SetPosition(8, 0, 0, 0);
                graph_per_change_70_22.SetSize(600, 440);
                ///////////////////////////////////////////////////////////////////////////////
                var graph_change_20_22 = change_20_22.Drawings.AddChart("Population change 2020-2022", eChartType.ColumnClustered3D);
                graph_change_20_22.SetPosition(0, 0, 10, 0);
                graph_change_20_22.SetSize(800, 600);

                var graph_per_change_20_22 = change_20_22.Drawings.AddChart("Population change (%) 2020-2022", eChartType.ColumnClustered);
                graph_per_change_20_22.SetPosition(8, 0, 0, 0);
                graph_per_change_20_22.SetSize(600, 440);
                ///////////////////////////////////////////////////////////////////////////////
                var grapg_per_world = percent_world.Drawings.AddChart("Population change (%) 2020-2022", eChartType.Pie3D);
                grapg_per_world.SetSize(600, 400);
                grapg_per_world.SetPosition(8, 0, 0, 0);
                grapg_per_world.Title.Text = "Percentage of population of each continent in the world population";
                ///////////////////////////////////////////////////////////////////////////////
                var graph_annual_asia = annual_change_70_22.Drawings.AddChart("Population change Asia 1970-2022", eChartType.LineMarkers);
                graph_annual_asia.SetPosition(8, 0, 0, 0);
                graph_annual_asia.SetSize(600, 440);
                graph_annual_asia.Title.Text = "Population change Asia 1970-2022";

                var graph_annual_africa = annual_change_70_22.Drawings.AddChart("Population change Africa 1970-2022", eChartType.LineMarkers);
                graph_annual_africa.SetPosition(31, 0, 0, 0);
                graph_annual_africa.SetSize(600, 440);
                graph_annual_africa.Title.Text = "Population change Africa 1970-2022";

                var graph_annual_SA = annual_change_70_22.Drawings.AddChart("Population change SouthAmerica 1970-2022", eChartType.LineMarkers);
                graph_annual_SA.SetPosition(0, 0, 10, 0);
                graph_annual_SA.SetSize(600, 440);
                graph_annual_SA.Title.Text = "Population change SouthAmerica 1970-2022";

                var graph_annual_NA = annual_change_70_22.Drawings.AddChart("Population change NorthAmerica 1970-2022", eChartType.LineMarkers);
                graph_annual_NA.SetPosition(23, 0, 10, 0);
                graph_annual_NA.SetSize(600, 440);
                graph_annual_NA.Title.Text = "Population change NorthAmerica 1970-2022";

                var graph_annual_europe = annual_change_70_22.Drawings.AddChart("Population change Europe 1970-2022", eChartType.LineMarkers);
                graph_annual_europe.SetPosition(0, 0, 20, 0);
                graph_annual_europe.SetSize(600, 440);
                graph_annual_europe.Title.Text = "Population change Europe 1970-2022";

                var graph_annual_oceania = annual_change_70_22.Drawings.AddChart("Population change Oceania 1970-2022", eChartType.LineMarkers);
                graph_annual_oceania.SetPosition(23, 0, 20, 0);
                graph_annual_oceania.SetSize(600, 440);
                graph_annual_oceania.Title.Text = "Population change Oceania 1970-2022";


                // Adding data to charts
                // worksheet.cells[x,y,x2,y2] - cells with data
                var seriesPopulation1970 = graph_change_70_22.Series.Add(change_70_22.Cells[2, 2, 7, 2], change_70_22.Cells[2, 1, 7, 1]);
                seriesPopulation1970.Header = "Population 1970";  //dodawanie nazwa seria
                var seriesPopulation2022 = graph_change_70_22.Series.Add(change_70_22.Cells[2, 3, 7, 3], change_70_22.Cells[2, 1, 7, 1]);
                seriesPopulation2022.Header = "Population 2022";

                var seriesPopulationChange1970To2022 = graph_per_change_70_22.Series.Add(change_70_22.Cells[2, 4, 7, 4], change_70_22.Cells[2, 1, 7, 1]);
                seriesPopulationChange1970To2022.Header = "Population Change (%) 1970-2022";
                /////////////////////////////////////////////////////////////////////
                var seriesPopulation2020 = graph_change_20_22.Series.Add(change_20_22.Cells[2, 2, 7, 2], change_20_22.Cells[2, 1, 7, 1]);
                seriesPopulation2020.Header = "Population 2020";
                var seriesPopulation2022_2 = graph_change_20_22.Series.Add(change_20_22.Cells[2, 3, 7, 3], change_20_22.Cells[2, 1, 7, 1]);
                seriesPopulation2022_2.Header = "Population 2022";

                var seriesPopulationChange2020To2022 = graph_per_change_20_22.Series.Add(change_20_22.Cells[2, 4, 7, 4], change_20_22.Cells[2, 1, 7, 1]);
                seriesPopulationChange2020To2022.Header = "Population Change (%) 2020-2022";
                ////////////////////////////////////////////////////////////////////////
                var seriesWorldSharePer2022 = grapg_per_world.Series.Add(percent_world.Cells[2, 3, continentName.Length + 1, 3], percent_world.Cells[2, 1, continentName.Length + 1, 1]);
                ////////////////////////////////////////////////////////////////////////
                var seriesPopulAsia = graph_annual_asia.Series.Add(annual_change_70_22.Cells[2, 2, 2, 9], annual_change_70_22.Cells[1, 2, 1, 9]);
                seriesPopulAsia.Header = "Population \n Asia";
                var seriesPopulAfrica = graph_annual_africa.Series.Add(annual_change_70_22.Cells[3, 2, 3, 9], annual_change_70_22.Cells[1, 2, 1, 9]);
                seriesPopulAfrica.Header = "Population \n Africa";
                var seriesPopulSA = graph_annual_SA.Series.Add(annual_change_70_22.Cells[4, 2, 4, 9], annual_change_70_22.Cells[1, 2, 1, 9]);
                seriesPopulSA.Header = "Population \n SouthAmerica";
                var seriesPopulNA = graph_annual_NA.Series.Add(annual_change_70_22.Cells[5, 2, 5, 9], annual_change_70_22.Cells[1, 2, 1, 9]);
                seriesPopulNA.Header = "Population \n NorthAmerica";
                var seriesPopulEurope = graph_annual_europe.Series.Add(annual_change_70_22.Cells[6, 2, 6, 9], annual_change_70_22.Cells[1, 2, 1, 9]);
                seriesPopulEurope.Header = "Population \n Europe";
                var seriesPopulOceania = graph_annual_oceania.Series.Add(annual_change_70_22.Cells[7, 2, 7, 9], annual_change_70_22.Cells[1, 2, 1, 9]);
                seriesPopulOceania.Header = "Population \n Oceania";


                // adjusting charts
                graph_change_70_22.Legend.Position = eLegendPosition.Right;
                graph_change_70_22.XAxis.Title.Text = "Continents";
                graph_change_70_22.XAxis.Title.Font.Bold = true;
                graph_change_70_22.YAxis.Title.Text = "Population Count";
                graph_change_70_22.YAxis.Title.Font.Bold = true;

                graph_per_change_70_22.XAxis.Title.Text = "Continents";
                graph_per_change_70_22.XAxis.Title.Font.Bold = true;
                graph_per_change_70_22.YAxis.Title.Text = "Population Change (%)";
                graph_per_change_70_22.YAxis.Title.Font.Bold = true;
                ///////////////////////////////////////////////////////////
                graph_change_20_22.Legend.Position = eLegendPosition.Right;
                graph_change_20_22.XAxis.Title.Text = "Continents";
                graph_change_20_22.XAxis.Title.Font.Bold = true;
                graph_change_20_22.YAxis.Title.Text = "Population Count";
                graph_change_20_22.YAxis.Title.Font.Bold = true;

                graph_per_change_20_22.XAxis.Title.Text = "Continents";
                graph_per_change_20_22.XAxis.Title.Font.Bold = true;
                graph_per_change_20_22.YAxis.Title.Text = "Population Change (%)";
                graph_per_change_20_22.YAxis.Title.Font.Bold = true;
                //////////////////////////////////////////////////////////


                // color setting
                seriesPopulation1970.Fill.Color = Color.LightBlue;
                seriesPopulation2022.Fill.Color = Color.LightGreen;
                seriesPopulationChange1970To2022.Fill.Color = Color.OrangeRed;
                ///////////////////////////////////////////////////////////
                seriesPopulation2020.Fill.Color = Color.LightBlue;
                seriesPopulation2022_2.Fill.Color = Color.LightGreen;
                seriesPopulationChange2020To2022.Fill.Color = Color.OrangeRed;
                //////////////////////////////////////////////////////////
                seriesPopulAsia.Border.Fill.Color = Color.DarkBlue;
                seriesPopulAfrica.Border.Fill.Color = Color.DarkBlue;
                seriesPopulSA.Border.Fill.Color = Color.DarkBlue;
                seriesPopulNA.Border.Fill.Color = Color.DarkBlue;
                seriesPopulEurope.Border.Fill.Color = Color.DarkBlue;
                seriesPopulOceania.Border.Fill.Color = Color.DarkBlue;


                package.SaveAs(new FileInfo(fileNameExcel));
            }
            catch (ArgumentException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (FileNotFoundException) { Console.WriteLine("\nError: The specified file to create the graphs does not exist!"); }
            catch (NullReferenceException) { Console.WriteLine("\nError : The Excel worksheet you are trying to reference is empty!"); }
            catch (InvalidOperationException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (IOException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (Exception ex) { Console.WriteLine("\nError : " + ex.Message); }
        }
        public void SaveToFile()
        {
            try
            {
                if (Path.GetExtension(fileNameToSave) != ".txt")
                {
                    throw new ArgumentException("Invalid file type. The file type must be .txt.");
                }

                StreamWriter sw = new StreamWriter(fileNameToSave, false);
                sw.WriteLine("continentName;continentPopulation1970;continentPopulation2022;changeInPopulation1970To2022(%)");
                for (int i = 0; i < continentName.Length; i++)
                {
                    sw.WriteLine("{0};{1};{2};{3}",
                        continentName[i],                     //Name of continent
                        continentPopulation1970To2022[i, 0],  //Population by continent 1970
                        continentPopulation1970To2022[i, 7],  //Population by continent 2022
                        changeInPopulation1970To2022[i]);
                }
                sw.WriteLine("\ncontinentName;continentPopulation2020;continentPopulation2022;changeInPopulation2020To2022(%)");
                for (int i = 0; i < continentName.Length; i++)
                {
                    sw.WriteLine("{0};{1};{2};{3}",
                        continentName[i],                     //Name of continent
                        continentPopulation1970To2022[i, 6],  //Population by continent 2020
                        continentPopulation1970To2022[i, 7],  //Population by continent 2022
                        changeInPopulation2020To2022[i]);
                }
                sw.WriteLine("worldPopulation2022 : " + worldPopulation2022);
                sw.WriteLine("\ncontinentName;continentPopulation2022;percentPopulationShareWorld(%)");
                for (int i = 0; i < continentName.Length; i++)
                {
                    sw.WriteLine("{0};{1};{2}",
                        continentName[i],                     //Name of continent
                        continentPopulation1970To2022[i, 6],  //Population by continent 2022
                        percentPopulationShareWorld[i]);      //Poppulation share world
                }
                sw.WriteLine("\ncontinentName;continentPop1970;pop1980;pop1990;pop2000;pop2010;pop2015;pop2020;pop2022");
                for (int i = 0; i < continentName.Length; i++)
                {
                    string line = $"{continentName[i]};";
                    for (int j = 0; j < continentPopulation1970To2022.GetLength(1); j++)
                    {
                        line += $"{continentPopulation1970To2022[i, j]};";
                    }
                    sw.WriteLine(line);
                }
                sw.Close();
            }
            catch (ArgumentException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (UnauthorizedAccessException) { Console.WriteLine("\nError : Does not have permission to write to the file!"); }
            catch (IOException e) { Console.WriteLine("\nError : " + e.Message); }
            catch (Exception ex) { Console.WriteLine("\nError : " + ex.Message); }
        }
    }
    class Program
    {
        static void Main()
        {
            Console.WriteLine("Enter the name of the file with the type of extension (e.g.txt) to read the data");
            string file_name_read = Console.ReadLine();

            Console.WriteLine("\nEnter the name of the Excel file with the type of extension (e.g.xlsx) to create the charts");
            string file_name_excel = Console.ReadLine();

            Console.WriteLine("\nEnter a file name with the type of extension (e.g.txt) for saving results.");
            string file_name_saving = Console.ReadLine();

            PopulationStatistics statistics = new PopulationStatistics(file_name_read, file_name_excel, file_name_saving);
            statistics.ReadFile();
            statistics.ContinentPopulationFrom1970to2022();
            statistics.PercentChangePopulation1970To2022();
            statistics.PercentChangePopulation2020To2022();
            statistics.PercentWorldPopulationShare2022();
            statistics.CreateChartsContinentPopulation();
            statistics.SaveToFile();
        }
    }
}
