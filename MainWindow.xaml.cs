using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Data.OleDb;
using System.Data;
using Plotly.NET.CSharp;
using Plotly.NET;
using Chart = Plotly.NET.CSharp.Chart;
using Plotly.NET.LayoutObjects;
using Microsoft.FSharp.Core;
using static Microsoft.FSharp.Core.ByRefKinds;
using System.Text.RegularExpressions;

namespace TADM_reader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public class TADMRecord(int Id, string LiquidClassName, string StepType, double Volume, int ChannelNumber, DateTime TimeStamp, List<Int16> CurvePoints)
    {
        public int Id = Id;
        public string LiquidClassName = LiquidClassName;
        public string StepType = StepType;
        public double Volume = Volume;
        public int ChannelNumber = ChannelNumber;
        public DateTime TimeStamp = TimeStamp;
        public List<Int16> CurvePoints = CurvePoints;
        public string PipettingStep = string.Empty; //The step corresponding to the curve: sample for counting, medium fill, seeding
        public string PipettingSource = string.Empty; //The well corresponding to the curve as reported in the trace file
        public override string ToString()
        {
            return $"{Id}, {LiquidClassName}, {StepType}, {Volume}, {ChannelNumber}, {TimeStamp}, {PipettingStep}, {PipettingSource}";
        }
        public string ToStringConcise()
        {
            return $"{Id}, {StepType}, {Volume}, {ChannelNumber}, {TimeStamp.TimeOfDay}, {PipettingStep}, {PipettingSource}";
        }
        public string CurvePointsToString()
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < CurvePoints.Count; i++)
            {
                sb.Append(CurvePoints[i]);
                if (i < CurvePoints.Count - 1)
                    sb.Append(";");
            }
            return sb.ToString();
        }

    }

    public static class filterClass
    {
        public static string liquidClassName = string.Empty;
        public static string stepType = string.Empty;
        public static double volumeMin = 0.0;
        public static double volumeMax = 5000.0;
        public static bool[] channel1mL = { true, true, true, true};
        public static bool[] channel5mL = { true, true};

    }
    public partial class filterWindow : Window
    {
        public filterWindow()
        {
            InitializeComponent();
        }

        private void textBoxVolumeMax_TextChanged(object sender, TextChangedEventArgs e)
        {

            filterClass.volumeMax = double.Parse(volumeMaxTextBox.Text);
            //volumeMinTextBox.Text = filterClass.volumeMax.ToString();
        }
        private void textBoxVolumeMin_TextChanged(object sender, TextChangedEventArgs e)
        {

            filterClass.volumeMin = double.Parse(volumeMinTextBox.Text);
            //volumeMinTextBox.Text = filterClass.volumeMax.ToString();
        }

        private void Chan1_1mL_Changed(object sender, RoutedEventArgs e)
        {
            filterClass.channel1mL[0] = (bool)Chan1_1mL_CheckBox.IsChecked;
        }

        private void Chan2_1mL_Changed(object sender, RoutedEventArgs e)
        {
            filterClass.channel1mL[1] = (bool)Chan2_1mL_CheckBox.IsChecked;
        }

        private void Chan3_1mL_Changed(object sender, RoutedEventArgs e)
        {
            filterClass.channel1mL[2] = (bool)Chan3_1mL_CheckBox.IsChecked;
        }
        private void Chan4_1mL_Changed(object sender, RoutedEventArgs e)
        {
            filterClass.channel1mL[3] = (bool)Chan4_1mL_CheckBox.IsChecked;
        }
        private void Chan1_5mL_Changed(object sender, RoutedEventArgs e)
        {
            filterClass.channel5mL[0] = (bool)Chan1_5mL_CheckBox.IsChecked;
        }
        private void Chan2_5mL_Changed(object sender, RoutedEventArgs e)
        {
            filterClass.channel5mL[1] = (bool)Chan2_5mL_CheckBox.IsChecked;
        }
        
    }

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Access Database Files (*.mdb;*.accdb)|*.mdb;*.accdb|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
                TextBlockFile.Text = openFileDialog.FileName;
            
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            TextBlock_log.Text = "";
            if (TextBlockFile.Text == string.Empty)
            {
                MessageBox.Show("Please select a file first.");
                TextBlock_log.Text += "No file selected\n";
                return;
            }
            try
            {
                var myDataTable = new DataTable();
                var oleDbconnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={TextBlockFile.Text};Persist Security Info=False;";
                //var oleDbconnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={TextBlockFile.Text};Persist Security Info=False;";

                TextBlock_log.Text += $"Attempting to open {TextBlockFile.Text} \n";

                using (var connection = new OleDbConnection(oleDbconnectionString))
                //using (var connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\caussems\\Documents\\ReadBinary\\Test.mdb;Persist Security Info=False;"))
                {
                    //Read database to retrieve all the curves points and metadata
                    connection.Open();
                    TextBlock_log.Text += "File opened, extracting data\n";
                    var query = "Select CurvePoints From TadmCurve  Order By CurveId";
                    var command = new OleDbCommand(query, connection);
                    var reader = command.ExecuteReader();
                    //string curvePoints = string.Empty;
                    var curvesPoints = new List<string>(); // Initialize as empty string to avoid null reference
                    var curvesPoints_int = new List<List<Int16>>();

                    var totalRecords = 0;
                    while (reader.Read())
                    {
                        curvesPoints.Add(string.Empty); // Add an empty string to initialize the list for the current record
                        curvesPoints.Add(Convert.ToHexString((byte[])reader[0]));
                        var curvePoint = Convert.ToHexString((byte[])reader[0]);
                        var curvePoints_int = new List<Int16>();
                        for (int i = 0; i < curvePoint.Length; i++)
                        //Read 2 by 2 to recontruct the original int values
                        {
                            string hexValue = curvePoint.Substring(i, 4);
                            byte[] bytes = Convert.FromHexString(hexValue);
                            Int16 value = BitConverter.ToInt16(bytes, 0);
                            curvePoints_int.Add(value);
                            //Console.WriteLine(value);
                            i += 3; // Move to the next 4 bytes (8 hex characters)
                        }
                        curvesPoints_int.Add(curvePoints_int);
                        totalRecords++;
                    }
                    Console.WriteLine($"Total records: {totalRecords}");
                    //Retrieves the metadata of the table to get the column names and types
                    //Retrieve CurveId
                    query = "Select CurveId From TadmCurve Order By CurveId";
                    command = new OleDbCommand(query, connection);
                    reader = command.ExecuteReader();
                    var CurvesId = new List<int>();
                    while (reader.Read())
                    {
                        CurvesId.Add((int)reader[0]);
                    }
                    //Retrieve StepType
                    query = "Select StepType From TadmCurve Order By CurveId";
                    command = new OleDbCommand(query, connection);
                    reader = command.ExecuteReader();
                    var StepType = new List<String>();
                    while (reader.Read())
                    {
                        if ((int)reader[0] == -533331728)
                            StepType.Add("Aspirate");
                        else
                            StepType.Add("Dispense"); // or any default value to indicate null
                    }
                    //Retrieve volume
                    query = "Select Volume From TadmCurve Order By CurveId";
                    command = new OleDbCommand(query, connection);
                    reader = command.ExecuteReader();
                    var volumes = new List<double>();
                    while (reader.Read())
                    {
                        volumes.Add((double)reader[0]);
                    }
                    //Retrieve TimeStamp
                    query = "Select TimeStamp From TadmCurve Order By CurveId";
                    command = new OleDbCommand(query, connection);
                    reader = command.ExecuteReader();
                    var TimeStamp = new List<DateTime>();
                    while (reader.Read())
                    {
                        TimeStamp.Add((DateTime)reader[0]);
                    }
                    //Retrieve Channel#
                    query = "Select ChannelNumber From TadmCurve Order By CurveId";
                    command = new OleDbCommand(query, connection);
                    reader = command.ExecuteReader();
                    var ChannelNumber = new List<int>();
                    while (reader.Read())
                    {
                        ChannelNumber.Add((int)reader[0]);
                    }
                    //Retrieve liquid class name
                    query = "Select LiquidClassName From TadmCurve Order By CurveId";
                    command = new OleDbCommand(query, connection);
                    reader = command.ExecuteReader();
                    var LiquidClassName = new List<String>();
                    while (reader.Read())
                    {
                        LiquidClassName.Add((String)reader[0]);
                    }
                    //Store all the records data in a list of TADMRecord objects to facilitate data manipulation and filtering
                    var allRecordList = new List<TADMRecord>();
                    var Records1mL = new List<TADMRecord>();
                    var Records5mL = new List<TADMRecord>();
                    for (int i = 0; i < totalRecords; i++)
                    {
                        allRecordList.Add(new TADMRecord(CurvesId[i], LiquidClassName[i], StepType[i], volumes[i], ChannelNumber[i], TimeStamp[i], curvesPoints_int[i]));
                        //allRecordList[i].PipettingSource = $"A{allRecordList[i].ChannelNumber}";
                        if (allRecordList[i].LiquidClassName.Contains("HighVolumeFilter") || allRecordList[i].LiquidClassName.Contains("1mL"))
                        {
                            Records1mL.Add(allRecordList[i]);
                        }
                        else
                        {
                            Records5mL.Add(allRecordList[i]);
                        }

                    }
                    //Try to infer which curves correspond to which well/step depending on the metadata and timing.
                    //Select only part of the curves based on filters : volume and liquid class name
                    TextBlock_log.Text += "Filtering data based on filter\n";
                    var recordFiltered = new List<TADMRecord>();
                    for (int i = 0; i<allRecordList.Count; i++)
                    {
                        if (allRecordList[i].Volume >= filterClass.volumeMin && allRecordList[i].Volume <= filterClass.volumeMax)
                        {
                            if (allRecordList[i].LiquidClassName.Contains("HighVolumeFilter") || allRecordList[i].LiquidClassName.Contains("1mL") || allRecordList[i].LiquidClassName.Contains("1000uL"))
                            {
                                if (filterClass.channel1mL[allRecordList[i].ChannelNumber - 1])
                                    recordFiltered.Add(allRecordList[i]);
                            }
                            else
                            {
                                if (filterClass.channel5mL[allRecordList[i].ChannelNumber - 1])
                                    recordFiltered.Add(allRecordList[i]);
                            }
                                
                        }
                        
                    }

                    //Build csv export file if option is checked
                    string saveFilePath = string.Empty;
                    if (CheckBoxCSV.IsChecked == true)
                    {
                        TextBlock_log.Text += "Creating csv output\n";
                        Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                        dlg.FileName = "output"; // Default file name
                        dlg.DefaultExt = ".csv"; // Default file extension
                        dlg.Filter = "Comma separated value (.csv)|*.csv"; // Filter files by extension

                        // Show save file dialog box
                        Nullable<bool> result = dlg.ShowDialog();

                        // Process save file dialog box results
                        if (result == true)
                        {
                            // Save document
                            StreamWriter writer = new StreamWriter(dlg.FileName);
                            //Write the header of the CSV file
                            writer.WriteLine("Id,LiquidClassName,StepType,Volume,Channel#,TimeStamp,CurvePoints");
                            for (int i = 0; i < recordFiltered.Count; i++)
                            {
                                writer.Write($"{recordFiltered[i].ToString()}, {recordFiltered[i].CurvePointsToString()}\n");
                            }
                            writer.Close();
                        }
                    }
                    //Try to find trace file file and find which curve correspond to which step/well
                    //remove ML_STAR_tadm.tadm from file path and replace by Trace.trc
                    var traceFilePath = TextBlockFile.Text.Replace("ML_STAR_tadm.mdb","Trace.trc");
                    //Parse the file and try to assign aspirate/dispense step to records
                    if(File.Exists(traceFilePath))
                    {
                        TextBlock_log.Text += $"Trace file found at {traceFilePath}\n";
                        StreamReader traceFile = new StreamReader(traceFilePath);
                        var recordIndex = 0;
                        while(!traceFile.EndOfStream)
                        {
                            var line = traceFile.ReadLine();
                            if ((line.Contains("1000ul Channel Aspirate (Single Step) - complete") || line.Contains("1000ul Channel Dispense (Single Step) - complete")))
                            {
                                //This line correspond to an aspirate or dispense step
                                //Find how many channels are concerned
                                var chanOccurences = Regex.Matches(line, "channel").Count;
                                var trimmedLine = line.Substring(line.IndexOf(";  >"));
                                //Due to the filter there should be more matching line than records. But the first record will match before any other, so we can just parse the file waiting for the match of the first record and only start incrementing from there
                                var recordIndexTemp = 0;
                                for (int i = recordIndex; i < recordIndex + chanOccurences; i++)
                                {
                                    if (i < recordFiltered.Count)
                                    {

                                        //As the log used different precision depending on the pipetting value apply a different truncating algo depending on value
                                        var truncateFactor = 100;
                                        if (recordFiltered[i].Volume > 100)
                                        {
                                            truncateFactor = 1;
                                        }
                                        var stringValue = (Math.Truncate(recordFiltered[i].Volume * truncateFactor) / truncateFactor).ToString();
                                        if (line.Contains(recordFiltered[i].StepType) && 
                                            (line.Contains(stringValue) || line.Contains(stringValue.Replace(".",",")) || line.Contains(stringValue.Replace(",",".")))
                                            && line.Contains($"channel {recordFiltered[i].ChannelNumber}"))
                                        {
                                            //Line is matching with the step type
                                            //Pipetting source is between the channel 1: and the second comma
                                            recordIndexTemp++;
                                            var indexStart = trimmedLine.IndexOf($"channel {recordFiltered[i].ChannelNumber}") + 10;
                                            
                                            var indexComma1 = trimmedLine.IndexOf(",",indexStart);
                                            var indexComma2 = trimmedLine.IndexOf(",", indexComma1 + 1);
                                            recordFiltered[i].PipettingSource = trimmedLine.Substring(indexStart, indexComma2-indexStart).Trim();
                                            //recordFiltered[i].PipettingSource = trimmedLine.Substring(trimmedLine.IndexOf($"channel {recordFiltered[i].ChannelNumber}") + 10, trimmedLine.IndexOf(",", trimmedLine.IndexOf(","))).Trim();

                                            if (recordFiltered[i].Volume == 170)
                                            {
                                                recordFiltered[i].PipettingStep = "Counting";
                                            }
                                            if (recordFiltered[i].Volume >= 1000)
                                            {
                                                recordFiltered[i].PipettingStep = "MediumFilling";
                                            }
                                            if (recordFiltered[i].PipettingSource.Contains("Medium"))
                                            {
                                                recordFiltered[i].PipettingStep = "MediumFilling";
                                            }
                                            if (recordFiltered[i].StepType.Contains("Dispense") && recordFiltered[i].PipettingSource.Contains("Child"))
                                            {
                                                //Find the previous record with same channel and Aspirate and set the same PipettingStep for this record
                                                for(int k = 1; k < 5; k++)
                                                {
                                                    if (i - k >= 0 && recordFiltered[i - k].ChannelNumber == recordFiltered[i].ChannelNumber && recordFiltered[i - k].StepType.Contains("Aspirate"))
                                                    {
                                                        recordFiltered[i].PipettingStep = recordFiltered[i - k].PipettingStep;
                                                        break;
                                                    }
                                                }
                                                //recordFiltered[i].PipettingStep = "Seeding"; //Or medium filling if next line contain Tip Eject ?
                                            }
                                            if (recordFiltered[i].StepType.Contains("Aspirate") && recordFiltered[i].PipettingSource.Contains("Parent") && recordFiltered[i].Volume != 170)
                                            {
                                                recordFiltered[i].PipettingStep = "Seeding";
                                            }
                                            if (recordFiltered[i].PipettingSource.Contains("banking"))
                                            {
                                                recordFiltered[i].PipettingStep = "Banking";
                                            }
                                            if (recordFiltered[i].Volume%250 == 0)
                                            {
                                                recordFiltered[i].PipettingStep = "Banking";
                                            }
                                        }
                                        
                                    }
                                }
                                recordIndex += recordIndexTemp;
                            }
                        }
                    }
                    else
                    {
                        TextBlock_log.Text += $"Trace file not found at {traceFilePath}\n";
                    }
                    
                    //Plot filtered curves using plotly, use metadata to set curve name and tooltip information
                    var curveMetadata = new List<string>();

                    TextBlock_log.Text += "Creating charts using plotly\n";
                    var chartListAspirate = new List<Plotly.NET.GenericChart>();
                    var chartListDispense = new List<Plotly.NET.GenericChart>();

                    var chartListAspirate5mL = new List<Plotly.NET.GenericChart>();
                    var chartListDispense5mL = new List<Plotly.NET.GenericChart>();
                    for (int i = 0; i < recordFiltered.Count; i++)
                    {
                        var curveLength = recordFiltered[i].CurvePoints.Count;
                        var x_length = new double[curveLength];
                        var y_length = new double[curveLength];
                        curveMetadata.Add(recordFiltered[i].ToStringConcise());
                        for (int j = 0; j < curveLength; j++)
                        {
                            x_length[j] = j * 10; //time in ms, 10 ms step
                            y_length[j] = recordFiltered[i].CurvePoints[j];
                        }

                        var lineChart = Chart.Line<double, double, string>(
                            x_length, y_length, Name: curveMetadata[i] //Use metadata to set curve name
                        );
                        if (recordFiltered[i].LiquidClassName.Contains("HighVolumeFilter") || recordFiltered[i].LiquidClassName.Contains("1mL") || recordFiltered[i].LiquidClassName.Contains("1000uL"))
                        {
                            if (curveMetadata[i].Contains("Aspirate"))
                            {
                                chartListAspirate.Add(lineChart);
                            }
                            else
                            {
                                chartListDispense.Add(lineChart);
                            }
                        }
                        else
                        {
                            if (curveMetadata[i].Contains("Aspirate"))
                            {
                                chartListAspirate5mL.Add(lineChart);
                            }
                            else
                            {
                                chartListDispense5mL.Add(lineChart);
                            }
                        }
                    }
                    var config = Config.init(Responsive: true, Autosizable: false, FillFrame: true);
                    LinearAxis xAxis = new LinearAxis();
                    xAxis.SetValue("title", "Time (ms)");
                    LinearAxis yAxis = new LinearAxis();
                    yAxis.SetValue("title", "Pressure (Pa)");
                    // Fix for CS0411: Specify the type argument explicitly for Layout.init
                    var screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
                    var screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
                    //TextBlock_log.Text += $"Screen width: {screenWidth}, Screen height: {screenHeight}\n";
                    var layout = Layout.init<bool>(AutoSize: (bool)true, Width: (int)(screenWidth*0.8), Height: (int)(screenHeight*0.8));
                    var titleAspirate = Plotly.NET.Title.init(Text: FSharpOption<string>.Some("Aspirate"));
                    var titleDispense = Plotly.NET.Title.init(Text: FSharpOption<string>.Some("Dispense"));
                    Legend legendAspirate = Legend.init(X: 1.02, Y: 1.1, Title: titleAspirate);

                    if (chartListAspirate.Count > 0)
                    {

                        // Fix for CS1503: Argument 3 and Argument 4 errors  
                        // The issue is that the arguments passed to `WithLegend` and `WithTraceInfo` methods are integers,  
                        // but the methods expect `FSharpOption<Plotly.NET.StyleParam.TraceItemClickOptions>` and  
                        // `FSharpOption<Plotly.NET.StyleParam.TraceGroupClickOptions>` respectively.  
                        // We need to wrap the integers in `FSharpOption.Some()` to match the expected type.  

                        var combinedChartAspirate = Chart.Combine(chartListAspirate)
                           .WithLayout(layout)
                           .WithConfig(config)
                           .WithXAxis(xAxis)
                           .WithYAxis(yAxis)
                           .WithLegend(Legend.init(
                               X: 1.02,
                               Y: 1.1,
                               Title: titleAspirate,
                               ItemClick: FSharpOption<Plotly.NET.StyleParam.TraceItemClickOptions>.Some(StyleParam.TraceItemClickOptions.Toggle),
                               GroupClick: FSharpOption<Plotly.NET.StyleParam.TraceGroupClickOptions>.Some(StyleParam.TraceGroupClickOptions.ToggleItem)
                           ));
                        /*.WithTraceInfo(
                            LegendGroup: FSharpOption<string>.Some("Group 1"),
                            LegendGroupTitle: titleAspirate
                        );*/
                        var combinedChartDispense = Chart.Combine(chartListDispense)
                           .WithLayout(layout)
                           .WithConfig(config)
                           .WithXAxis(xAxis)
                           .WithYAxis(yAxis)
                           .WithLegend(Legend.init(
                               X: 1.02,
                               Y: 0.5,
                               Title: titleDispense,
                               ItemClick: FSharpOption<Plotly.NET.StyleParam.TraceItemClickOptions>.Some(StyleParam.TraceItemClickOptions.Toggle),
                               GroupClick: FSharpOption<Plotly.NET.StyleParam.TraceGroupClickOptions>.Some(StyleParam.TraceGroupClickOptions.ToggleItem)
                           ));
                           /*.WithTraceInfo(
                               LegendGroup: FSharpOption<string>.Some("Group 2"),
                               LegendGroupTitle: titleDispense
                           );*/
                        //var combinedChartAspirate = Chart.Combine(chartListAspirate).WithLayout(layout).WithConfig(config).WithXAxis(xAxis).WithYAxis(yAxis).WithLegend(Legend.init(Title: titleAspirate)).WithTraceInfo(LegendGroup: FSharpOption<string>.Some("Group 1"), LegendGroupTitle: titleAspirate);
                        //var combinedChartDispense = Chart.Combine(chartListDispense).WithLayout(layout).WithConfig(config).WithXAxis(xAxis).WithYAxis(yAxis).WithLegend(Legend.init(Title: titleDispense)).WithTraceInfo(LegendGroup: FSharpOption<string>.Some("Group 2"), LegendGroupTitle: titleDispense);
                        var combinedAll = new List<Plotly.NET.GenericChart>() { combinedChartAspirate, combinedChartDispense };
                        var plotTitles = new List<string>() { "Aspirate 1mL", "Dispense 1mL" };


                        var fullPlot = Chart.Grid(combinedAll, 2, 1, SubPlotTitles: plotTitles).WithTitle($"TADM results from process ran on {recordFiltered[0].TimeStamp.Date.ToShortDateString()} Filter applied :{filterClass.volumeMin} < volume > {filterClass.volumeMax}" +
                            $" Channel 1mL :{filterClass.channel1mL[0]} {filterClass.channel1mL[1]} {filterClass.channel1mL[2]} {filterClass.channel1mL[3]} ")
                            .WithLayout(layout);
                            
                        Plotly.NET.CSharp.GenericChartExtensions.Show(fullPlot);
                    }
                    if (chartListAspirate5mL.Count > 0)
                    {
                        var combinedChartAspirate5mL = Chart.Combine(chartListAspirate5mL).WithConfig(config).WithLayout(layout).WithXAxis(xAxis).WithYAxis(yAxis).WithLegend(Legend.init(X: 1.02, Y: 1.1))
                            .WithTraceInfo(LegendGroup: FSharpOption<string>.Some("Group 1"), LegendGroupTitle: titleAspirate); ;
                        var combinedChartDispense5mL = Chart.Combine(chartListDispense5mL).WithConfig(config).WithLayout(layout).WithXAxis(xAxis).WithYAxis(yAxis).WithLegend(Legend.init(X: 1.02, Y: 0.5))
                            .WithTraceInfo(LegendGroup: FSharpOption<string>.Some("Group 2"), LegendGroupTitle: titleDispense);
                        var combinedAll5mL = new List<Plotly.NET.GenericChart>() { combinedChartAspirate5mL, combinedChartDispense5mL };
                        var plotTitles5mL = new List<string>() { "Aspirate 5mL", "Dispense 5mL" };

                        var fullPlot5mL = Chart.Grid(combinedAll5mL, 2, 1, SubPlotTitles: plotTitles5mL).WithTitle($"TADM results from process ran on {recordFiltered[0].TimeStamp.Date.ToShortDateString()} Filter applied :{filterClass.volumeMin} < volume > {filterClass.volumeMax}," +
                            $" Channel 5mL : {filterClass.channel5mL[0]} {filterClass.channel5mL[1]}").WithLayout(layout);
                        Plotly.NET.CSharp.GenericChartExtensions.Show(fullPlot5mL);
                    }
                    TextBlock_log.Text += "Successfully created the plots.\n";
                    //TODO : Select a group of curves depending on the metadata and only plot those curves.
                    //add the possibility to select multiple files and plot them all together
                    //select curves between 2 volumes
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error running the application: {ex.Message}");
                TextBlock_log.Text += $"Error running the application: {ex.Message} \n";
                return;
            }
            
         
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var filterWindow = new filterWindow();

            filterWindow.Owner = this;
            filterWindow.Show();
        }
    }
}