// -*-C++-*-
/*!
 * @file  myExcel.h
 *
 */
#ifndef MYEXCEL_H
#define MYEXCEL_H

using namespace Microsoft::Office::Interop::Excel;
#define Excel   Microsoft::Office::Interop::Excel

#include <iostream>
#include <string>
#include <vector>

#include "MyObject.h"



/**
* @class myExcel
*@brief Excelを操作するクラス
*/

ref class myExcel
{
public:
	/**
	*@brief コンストラクタ
	*/
	myExcel();
	static myExcel^ Obj;	/**<　@brief  */
	
	/**
	*@brief 名前からワークシートを取得する関数
	* @param sn
	* @return 
	*/
	Excel::Worksheet^ GetWorksheet(std::string sn);
	
	/**
	*@brief 保存用ワークシートに書き込む関数
	* @param sf
	*/
	void SaveRTC(std::vector<std::string> sf);
	
	/**
	*@brief データポートの情報を保存用ワークシートから読み込む関数
	* @return 
	*/
	std::vector<std::string> LoadRTC();
	
	/**
	*@brief セルの色を無色にする関数
	* @param c
	* @param l
	* @param sn
	* @param leng
	*/
	void ResetCellColor(int c, std::string l, std::string sn, std::string leng);
	

	/**
	*@brief 単一のセルの文字列を変更する関数
	* @param c
	* @param l
	* @param sn
	* @param val
	*/
	
	void SetCellStringSingle(int c, std::string l, std::string sn, std::string val);
	
	/**
	*@brief 単一のセルの値を変更する関数
	* @param c
	* @param l
	* @param sn
	* @param val
	*/
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

	
	/**
	*@brief 単一のセルの値を取得する関数
	* @param c
	* @param l
	* @param sn
	* @return 
	*/
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
	
	/**
	*@brief 1つ以上のセルの値を変更する関数
	* @param c
	* @param l
	* @param sn
	* @param state
	* @param v
	*/
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

	
	
	
	/**
	*@brief 1つ以上のセルの値を取得する関数
	* @param c
	* @param l
	* @param sn
	* @param state
	* @return 
	*/
	template <typename T>
	std::vector<T> GetCellValue(int c, std::string l, std::string sn, std::string leng)
	{
		Excel::Worksheet^ws = GetWorksheet(sn);

		std::vector<T> v;

		

		if(ws != nullptr)
		{
			int t_l = convertStrToVal(l);
			int t_leng = t_l;
			
			try{
				t_leng = convertStrToVal(leng);		
			}
			catch(...)
			{
				t_leng = t_l;
			}
			if(t_l > t_leng)
				t_leng = t_l;
			
			

			

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
	
	/**
	*@brief Excelファイルを開く関数
	* @param fn
	*/
	void Open(System::String^ fn);
	
	/**
	*@brief COMオブジェクトを解放する関数
	*/
	void Close();

	/**
	*@brief
	* @param r
	* @param g
	* @param b
	*/
	void SetColor(int r, int g, int b);

	int red, green, blue; /**<　@brief  */

	System::String^ filename; /**<　@brief  */
	Excel::Application^ xlApplication; /**<　@brief  */
	Excel::Workbooks^ xlWorkbooks; /**<　@brief  */
	Excel::Workbook^ xlWorkbook; /**<　@brief  */
	Excel::Sheets^ xlWorksheets; /**<　@brief  */

	System::Collections::Generic::List<Excel::Worksheet^>^ xlWorksheet; /**<　@brief  */

};

#endif