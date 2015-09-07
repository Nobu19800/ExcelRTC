/*!
 * @file  ExcelObject.cpp
 * @brief Excelの操作
 *
 */

#include "stdafx.h"
#include "ExcelObject.h"
#include "SubFunction.h"
#include <iostream>


/**
*@brief Excelを操作するクラスのコンストラクタ
*/
ExcelObject::ExcelObject()
{
	xlApplication = nullptr;
	xlWorkbooks = nullptr;
	xlWorkbook = nullptr;
	xlWorksheets = nullptr;

	xlWorksheet = gcnew System::Collections::Generic::List<Excel::Worksheet^>();
	filename = gcnew System::String(" ");
	//Open("");

	red = 255;
	green = 255;
	blue = 0;
}

/**
*@brief データ取得中のセルの背景の色を設定
* @param r 背景の色(R)
* @param g 背景の色(G)
* @param b 背景の色(B)
*/
void ExcelObject::setColor(int r, int g, int b)
{
	red = r;
	green = g;
	blue = b;
}

/**
*@brief 名前からワークシートを取得する関数
* @param sn シート名
* @return シートオブジェクト
*/
Excel::Worksheet^ ExcelObject::getWorksheet(std::string sn)
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

/**
*@brief セルの色を無色にする関数
* @param c 列番号
* @param l 行番号
* @param sn シート名
* @param leng 行の範囲
*/
void ExcelObject::resetCellColor(int c, std::string l, std::string sn, std::string len)
{
	Excel::Worksheet^ws = getWorksheet(sn);
	if(ws != nullptr)
	{
		int t_l = convertStrToVal(l);
		int t_leng = convertStrToVal(len);


		if(c > 0 && t_l > 0 && t_leng >= t_l)
		{
			Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,t_l]);
			Excel::Range^c2 = safe_cast<Excel::Range^>(ws->Cells[c,t_leng]);
			
			
			Excel::Range^mr = ws->Range[c1,c2];

			mr->Interior->ColorIndex = 0;
		}
		

	}
}

/**
*@brief 単一のセルの文字列を変更する関数
* @param c 列番号
* @param l 行番号
* @param sn シート名
* @param val 入力する文字列
*/
void ExcelObject::setCellStringSingle(int c, std::string l, std::string sn, std::string val)
{
	int t_l = convertStrToVal(l);
	Excel::Worksheet^ws = getWorksheet(sn);
	if(ws != nullptr)
	{
		if(c > 0 && t_l > 0)
		{
			Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,t_l]);
			c1->Value2 = gcnew System::String(val.c_str());
		}
	}
}

/**
*@brief 保存用ワークシートに書き込む関数
* @param sf 書き込む文字列
*/
void ExcelObject::saveRTC(std::vector<std::string> sf)
{
	Excel::Worksheet^ws = getWorksheet("保存用");

	if(ws != nullptr)
	{
		for(int i=0;i < sf.size();i++)
		{
			Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[1+i,1]);
			c1->Value2 = gcnew System::String(sf[i].c_str());
		}
	}
}

/**
*@brief データポートの情報を保存用ワークシートから読み込む関数
* @return 読み込んだ文字列
*/
std::vector<std::string> ExcelObject::loadRTC()
{
	Excel::Worksheet^ws = getWorksheet("保存用");
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


/**
*@brief Excelファイルを開く関数
* @param fn ファイル名
*/
void ExcelObject::Open(System::String^ fn)
{
	
	if(filename == fn)
	{
		return;
	}

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
	 
	 Excel::Worksheet^ws = getWorksheet("保存用");
	 if(ws == nullptr)
	 {
		 xlWorksheets->Add(System::Reflection::Missing::Value,xlWorksheet[xlWorksheet->Count-1],System::Reflection::Missing::Value,System::Reflection::Missing::Value);
		 Worksheet^ wsp = static_cast<Worksheet^>(xlWorksheets->Item[xlWorksheets->Count]);
		 wsp->Name = "保存用";
		 xlWorksheet->Add(wsp);
		 xlWorksheets->Select(System::Reflection::Missing::Value);

		 //Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[1,1]);
		 //c1->Value2 = gcnew System::String("a");
		 
		 //xlWorksheets->Add("保存用",System::Reflection::Missing::Value,System::Reflection::Missing::Value,System::Reflection::Missing::Value);
	 }
	 else
	 {
		
	 }

	
			
}

/**
*@brief COMオブジェクトを解放する関数
*/
void ExcelObject::Close()
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