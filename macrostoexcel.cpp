// macrostoexcel.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

// ATL
#include <atlbase.h>

// Office
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE14\\MSO.DLL" \
    rename("RGB", "MSORGB") \
    rename("DocumentProperties", "MSODocumentProperties")
using namespace Office;

// VB
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB"
using namespace VBIDE;

//  Excel
#import "C:\\Program Files\\Microsoft Office\\Office14\\EXCEL.EXE" \
    rename("DialogBox", "ExcelDialogBox") \
    rename("RGB", "ExcelRGB") \
    rename("CopyFile", "ExcelCopyFile") \
    rename("ReplaceText", "ExcelReplaceText") \
    no_auto_exclude
using namespace Excel;

// STL
#include <iostream>
#include <string>

using namespace std;
 
char* get_cmd_option(char ** begin, char ** end, const std::string & option)
{
	auto itr = std::find(begin, end, option);

	if (itr != end && ++itr != end)
	{
		return *itr;
	}

	return nullptr;
}

int main(const int argc, char** argv)
{
	const auto source_file = get_cmd_option(argv, argv + argc, "-s");
	const auto dest_file =  get_cmd_option(argv, argv + argc, "-d");

	if (!source_file || !dest_file)
	{
		cout << "No argument provided!" << endl;

		return -1;
	}

	cout << "Source file: " << source_file << endl;
	cout << "Dest file: " << dest_file << endl;

	CoInitialize(nullptr);

	cout << "COM Initialize" << endl;

	_ApplicationPtr excel;
	int finished;

	try
	{
		//init options
		auto hr = excel.GetActiveObject(_T("Excel.Application"));
		if (FAILED(hr))
		{
			hr = excel.CreateInstance(_T("Excel.Application"), nullptr, CLSCTX_ALL);

			if (FAILED(hr))
			{
				throw hr;
			}
		}
		cout << "Excel Application is started" << endl;
		excel->PutVisible(0, VARIANT_FALSE);
		excel->PutUserControl(VARIANT_FALSE);
		excel->PutDisplayAlerts(0, VARIANT_FALSE);
		excel->PutShowWindowsInTaskbar(VARIANT_FALSE);

		//open opetions
		auto work_book = excel->Workbooks->Open(source_file);
		excel->PutVisible(0, VARIANT_FALSE);
		excel->PutUserControl(VARIANT_FALSE);
		excel->PutDisplayAlerts(0, VARIANT_FALSE);
		excel->PutShowWindowsInTaskbar(VARIANT_FALSE);
		const auto module_name = "your_init_macros_method_name";
		excel->Run(module_name);

		//save options
		excel->PutVisible(0, VARIANT_FALSE);
		excel->PutUserControl(VARIANT_FALSE);
		excel->PutDisplayAlerts(0, VARIANT_FALSE);
		excel->PutShowWindowsInTaskbar(VARIANT_FALSE);
		work_book->SaveAs(dest_file, xlOpenXMLWorkbook, "", "", NULL, NULL, xlExclusive);

		// Close the workbook
		work_book->Close(xlDoNotSaveChanges);

		// Quit the Excel application. (i.e. Application.Quit)
		excel->Quit();
		cout << "Quit the Excel application" << endl;

		// Release the COM objects.
		// Releasing the references is not necessary for the smart pointers
		excel->Release();

		finished = 0;
	}
	catch (_com_error &err)
	{
		cout << "Excel throws the error: " << err.ErrorMessage() << endl;
		cout << "Description: " << err.Description() << "Help file - " << err.HelpFile() << "Source- " << err.Source() << endl;

		finished = -1;
	}

	CoUninitialize();
	cout << "COM Un-initialize" << endl;

	return finished;
}

