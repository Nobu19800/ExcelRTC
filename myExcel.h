#ifndef MYEXCEL_H
#define MYEXCEL_H

using namespace Microsoft::Office::Interop::Excel;
#define Excel   Microsoft::Office::Interop::Excel

#include <iostream>
#include <string>
#include <vector>

#include "MyObject.h"


//Excel�𑀍삷��N���X

ref class myExcel
{
public:
	myExcel();
	//���O���烏�[�N�V�[�g���擾����֐�
	Excel::Worksheet^ GetWorksheet(std::string sn);
	//�ۑ��p���[�N�V�[�g�ɏ������ފ֐�
	void SaveRTC(std::vector<std::string> sf);
	//�f�[�^�|�[�g�̏���ۑ��p���[�N�V�[�g����ǂݍ��ފ֐�
	std::vector<std::string> LoadRTC();
	//�Z���̐F�𖳐F�ɂ���֐�
	void ResetCellColor(int c, std::string l, std::string sn, std::string leng);
	
	//�P��̃Z���̒l��ύX����֐�
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

	//�P��̃Z���̒l���擾����֐�
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
	//1�ȏ�̃Z���̒l��ύX����֐�
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

	
	
	//1�ȏ�̃Z���̒l���擾����֐�
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
	//Excel�t�@�C�����J���֐�
	void Open(System::String^ fn);
	//COM�I�u�W�F�N�g���������֐�
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