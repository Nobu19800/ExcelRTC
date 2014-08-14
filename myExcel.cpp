#include "stdafx.h"
#include "myExcel.h"
#include "MyObject.h"
#include <iostream>


myExcel::myExcel()
{
	xlApplication = nullptr;
	xlWorkbooks = nullptr;
	xlWorkbook = nullptr;
	xlWorksheets = nullptr;

	xlWorksheet = gcnew System::Collections::Generic::List<Excel::Worksheet^>();

	Open("");
}

Excel::Worksheet^ myExcel::GetWorksheet(std::string sn)
{
	for(int i=0;i < xlWorksheet->Count;i++)
	{
		std::string ts = MarshalString(xlWorksheet[i]->Name);
		
		if(ts == sn)
		{
			return xlWorksheet[i];
		}
	}
	return nullptr;
}

void myExcel::ResetCellColor(int c, std::string l, std::string sn, std::string len)
{
	Excel::Worksheet^ws = GetWorksheet(sn);
	if(ws != nullptr)
	{
		int t_l = convertStrToVal(l);
		int t_leng = convertStrToVal(len);


		Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,t_l]);
		Excel::Range^c2 = safe_cast<Excel::Range^>(ws->Cells[c,t_leng]);
		
		
		Excel::Range^mr = ws->Range[c1,c2];

		mr->Interior->ColorIndex = 0;

	}
}

void myExcel::SaveRTC(std::vector<std::string> sf)
{
	Excel::Worksheet^ws = GetWorksheet("保存用");

	if(ws != nullptr)
	{
		for(int i=0;i < sf.size();i++)
		{
			Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[1+i,1]);
			c1->Value2 = gcnew System::String(sf[i].c_str());
		}
	}
}
std::vector<std::string> myExcel::LoadRTC()
{
	Excel::Worksheet^ws = GetWorksheet("保存用");
	std::vector<std::string> v;

	if(ws != nullptr)
	{
		for(int i=0;i < 100;i++)
		{
			Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[1+i,1]);
			//c1->Value2 = gcnew System::String("");
			//c1->Text->ToString();
			try{
				std::string tmp = MarshalString(c1->Text->ToString());
				if(tmp == "")
				{
					return v;
				}
				else
				{
					v.push_back(tmp);
				}
			}
			catch(...)
			{
				return v;
			}
		}
		
	}
	return v;
}



void myExcel::Open(System::String^ fn)
{
	Excel::Application^ t_xlApplication;
	Excel::Workbooks^ t_xlWorkbooks;
	Excel::Workbook^ t_xlWorkbook;

	
	filename = fn;
	 try {
		 try {
			System::Object^tObject = System::Runtime::InteropServices::Marshal::GetActiveObject("Excel.Application");
			t_xlApplication = static_cast<Excel::Application^>(tObject);
			}
		 catch(...){
				t_xlApplication = gcnew Excel::Application();
			}
			t_xlApplication->Visible = true;
			try {
				t_xlWorkbooks = t_xlApplication->Workbooks;
				

				try {
					if(filename == "")
					{
						t_xlWorkbook = t_xlWorkbooks->Add(System::Reflection::Missing::Value);
					}
					else
					{
						t_xlWorkbook = t_xlWorkbooks->Open(filename, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value, System::Reflection::Missing::Value);
					}
					
					try{
						Close();
						xlApplication = t_xlApplication;
						xlWorkbooks = t_xlWorkbooks;
						xlWorkbook = t_xlWorkbook;


						xlWorksheets = xlWorkbook->Worksheets;
						
						int count = xlWorksheets->Count;

						for(int i=1;i <= count;i++)
						{
							xlWorksheet->Add(static_cast<Worksheet^>(xlWorksheets->Item[i]));
						}
						
					}
					catch(...)
					{

					}
				}
				catch(...)
				{

				}
			}
			catch(...)
			{

			}
	 }
	 catch(...)
	 {

	 }

	 Excel::Worksheet^ws = GetWorksheet("保存用");
	 if(ws == nullptr)
	 {
		 xlWorksheets->Add(System::Reflection::Missing::Value,xlWorksheet[xlWorksheet->Count-1],System::Reflection::Missing::Value,System::Reflection::Missing::Value);
		 Worksheet^ ws = static_cast<Worksheet^>(xlWorksheets->Item[xlWorksheets->Count]);
		 ws->Name = "保存用";
		 xlWorksheet->Add(ws);
		 xlWorksheets->Select(System::Reflection::Missing::Value);

		 //Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[1,1]);
		 //c1->Value2 = gcnew System::String("a");
		 
		 //xlWorksheets->Add("保存用",System::Reflection::Missing::Value,System::Reflection::Missing::Value,System::Reflection::Missing::Value);
	 }
	 else
	 {
		
	 }

	
			
}

void myExcel::Close()
{
	int count = xlWorksheet->Count;
	for(int i=1;i < count;i++)
	{
		if (xlWorksheet[i] != nullptr) {
			System::Runtime::InteropServices::Marshal::ReleaseComObject(xlWorksheet[i]);
		}
	}
	xlWorksheet->Clear();
	if (xlWorksheets != nullptr) {
		System::Runtime::InteropServices::Marshal::ReleaseComObject(xlWorksheets);
	}
	xlWorksheets = nullptr;

	if (xlWorkbook != nullptr) {
		System::Runtime::InteropServices::Marshal::ReleaseComObject(xlWorkbook);
	}
	xlWorkbook = nullptr;

	if (xlWorkbooks != nullptr) {
		System::Runtime::InteropServices::Marshal::ReleaseComObject(xlWorkbooks);
	}
	xlWorkbooks = nullptr;

	if (xlApplication != nullptr) {
		System::Runtime::InteropServices::Marshal::ReleaseComObject(xlApplication);
	}
	xlApplication = nullptr;
}