namespace SampleApp;

internal class Program
{
    static void Main(string[] args)
    {
        string filePath = "sampledocs-50mb-xlsx-file-sst.xlsx";//"sampledocs-50mb-xlsx-file.xlsx";

        //// To run samples un comment below code
        //Samples.Sample1(filePath);
        //Samples.Sample2(filePath);

        //To check comparison uncomment below line 1 by 1 
        Comparison.XlsxHelper(filePath);
        //Comparison.ExcelDataReader(filePath);
        //Comparison.ExcelDataReaderAsDataset(filePath);
        //Comparison.LightweightExcelReader(filePath);
    }
}