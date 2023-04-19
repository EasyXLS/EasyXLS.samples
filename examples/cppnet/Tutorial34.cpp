/*------------------------------------------------------------------------
 | Tutorial 34
 |
 | This tutorial shows how to import Excel to DataSet in C++.NET. The data
 | is imported from the active sheet of the Excel file (the Excel file
 | generated in Tutorial 09).
 -------------------------------------------------------------------------*/

using namespace System;
using namespace System::IO;
using namespace System::Data;
using namespace EasyXLS;
using namespace System::Text;

int main()
{
    Console::WriteLine("Tutorial 34\n----------\n");

    // Create an instance of the class that imports Excel files
    ExcelDocument ^workbook = gcnew ExcelDocument();

    // Import Excel file to DataSet
    Console::WriteLine("Reading file C:\\Samples\\Tutorial09.xlsx.\n");
    DataSet ^ds = workbook->easy_ReadXLSXActiveSheet_AsDataSet("C:\\Samples\\Tutorial09.xlsx");

    // Confirm import of Excel file
    String ^sError = workbook->easy_getError();
    if (sError->Equals(""))
    {
        // Display imported DataSet values
        DataTable ^dt = ds->Tables[0];
        StringBuilder ^str;
        for (int row=0; row < dt->Rows->Count; row++)
        {
            for (int column=0; column < dt->Columns->Count; column++)
            {
                str = gcnew StringBuilder();
                str->Append(String::Concat("At row ", (row + 1).ToString(), ", column ", (column + 1).ToString()));
                str->Append(String::Concat(" the value is '", 
                    Convert::ToString(dt->Rows[row]->ItemArray[column]), "'"));
                Console::WriteLine(str);
            }
        }
    }
    else
        Console::Write(String::Concat("\nError reading file C:\\Samples\\Tutorial09.xls \n", sError));

    // Dispose memory
    delete workbook;

    Console::Write("\nPress Enter to Exit...");
    Console::ReadLine();

    return 0;
}