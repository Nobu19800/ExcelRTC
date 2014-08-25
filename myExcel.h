#ifndef MYEXCEL_H
#define MYEXCEL_H

using namespace Microsoft::Office::Interop::Excel;
#define Excel   Microsoft::Office::Interop::Excel

#include <iostream>
#include <string>
#include <vector>

#include "MyObject.h"


//Excelを操作するクラス

ref class myExcel
{
public:
	myExcel();
	//名前からワークシートを取得する関数
	Excel::Worksheet^ GetWorksheet(std::string sn);
	//保存用ワークシートに書き込む関数
	void SaveRTC(std::vector<std::string> sf);
	//データポートの情報を保存用ワークシートから読み込む関数
	std::vector<std::string> LoadRTC();
	//セルの色を無色にする関数
	void ResetCellColor(int c, std::string l, std::string sn, std::string leng);
	
	//単一のセルの値を変更する関数
	template <typename T>
	void SetCellValueSingle(int c, std::string l, std::string sn, T val)
	{
		int t_l = convertStrToVal(l);
		Excel::Worksheet^ws = GetWorksheet(sn);
		if(ws != nullptr)
		{
			if(c > 0 && t_l > 0)
			{
				Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,t_l]);
				c1->Value2 = val;
			}
		}
	}

	//単一のセルの値を取得する関数
	std::string GetCellValueSingle(int c, std::string l, std::string sn)
	{
		int t_l = convertStrToVal(l);
		Excel::Worksheet^ws = GetWorksheet(sn);
		if(ws != nullptr)
		{
			if(c > 0 && t_l > 0)
			{
				Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,t_l]);
			
				return MarshalString(c1->Value2->ToString());
			}
		}
		return "error";
	}
	//1つ以上のセルの値を変更する関数
	template <typename T>
	void SetCellValue(int c, std::string l, std::string sn, bool state, std::vector<std::vector<T>> v)
	{
		Excel::Worksheet^ws = GetWorksheet(sn);

		
		if(ws != nullptr)
		{
			int t_l = convertStrToVal(l);

			int mnum = v.size();
			if(!state)
				mnum = 1;
			for(int i=0;i < mnum;i++)
			{
				if(c+i > 0 && t_l > 0)
				{
					Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c+i,t_l]);
					Excel::Range^c2 = safe_cast<Excel::Range^>(ws->Cells[c+i,t_l+v[i].size()-1]);
					
					
					Excel::Range^mr = ws->Range[c1,c2];		
					
					array<T>^dt = gcnew array<T>(v[i].size());
					//std::cout << v.size() << std::endl;
					for(int j=0;j < v[i].size();j++)
					{
						
						dt[j] = v[i][j];//gcnew System::String(v[i].c_str());
						
						//mr->Value2 = gcnew System::String(v.c_str());
						//ws->Cells[c,l] =  gcnew System::String(v.c_str());
					}

					
					
					mr->Value2 = dt;
				}
			}

			
			//System::Runtime::InteropServices::Marshal::ReleaseComObject(c1);
			//System::Runtime::InteropServices::Marshal::ReleaseComObject(c2);
			//System::Runtime::InteropServices::Marshal::ReleaseComObject(mr);
			
		}
	}

	
	
	//1つ以上のセルの値を取得する関数
	template <typename T>
	std::vector<T> GetCellValue(int c, std::string l, std::string sn, std::string leng)
	{
		Excel::Worksheet^ws = GetWorksheet(sn);

		std::vector<T> v;

		

		if(ws != nullptr)
		{
			int t_l = convertStrToVal(l);
			int t_leng = convertStrToVal(leng);

			

			if(c > 0 && t_l > 0 && t_leng >= t_l)
			{
				

				if(c > 1){
					

					Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c-1,t_l]);
					Excel::Range^c2 = safe_cast<Excel::Range^>(ws->Cells[c-1,t_leng]);
					
					
					Excel::Range^mr = ws->Range[c1,c2];

					mr->Interior->ColorIndex = 0;
				}

				{
					Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,t_l]);
					Excel::Range^c2 = safe_cast<Excel::Range^>(ws->Cells[c,t_leng]);
					
					
					Excel::Range^mr = ws->Range[c1,c2];

					mr->Interior->ColorIndex = 6;
					//mr->Interior->Color = System::Drawing::Color::FromArgb(red, blue, green);
					
				}
			}

			//Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,l]);
			//Excel::Range^c2 = safe_cast<Excel::Range^>(ws->Cells[c,l+leng-1]);
			
			
			//Excel::Range^mr = ws->Range[c1,c2];	


			
			for(int i=0;i < t_leng-t_l+1;i++)
			{
				if(c > 0 && t_l+i > 0)
				{
					Excel::Range^c1 = safe_cast<Excel::Range^>(ws->Cells[c,t_l+i]);
					/*if(c > 1)
					{
						Excel::Range^c2 = safe_cast<Excel::Range^>(ws->Cells[c-1,l+i]);
						c2->Interior->ColorIndex = 0;
					}
					c1->Interior->ColorIndex = 6;*/

					try{
						
						
						v.push_back(string2binary<T>(MarshalString(c1->Value2->ToString()),10));
						
					}
					catch(...)
					{
						v.push_back(0);
					}
				}
				//Excel::Range^mr2 = safe_cast<Excel::Range^>(mr[0,i]);
				//v.push_back(string2binary<T>(MarshalString(mr[0,i]->ToString()),10));
				//v.push_back(MarshalString(mr2->Value2->ToString()));
				//System::Runtime::InteropServices::Marshal::ReleaseComObject(mr2);
			}

			
			//mr->Text->ToString();
			
			
			//System::Runtime::InteropServices::Marshal::ReleaseComObject(mr);
			//System::Runtime::InteropServices::Marshal::ReleaseComObject(c1);
			//System::Runtime::InteropServices::Marshal::ReleaseComObject(c2);
			
			
		}

		return v;
	}
	//Excelファイルを開く関数
	void Open(System::String^ fn);
	//COMオブジェクトを解放する関数
	void Close();

	void SetColor(int r, int g, int b);

	int red, green, blue;

	System::String^ filename;
	Excel::Application^ xlApplication;
	Excel::Workbooks^ xlWorkbooks;
	Excel::Workbook^ xlWorkbook;
	Excel::Sheets^ xlWorksheets;

	System::Collections::Generic::List<Excel::Worksheet^>^ xlWorksheet;

};

#endif