// -*-C++-*-
/*!
 * @file  SpreadSheetSVC_impl.cpp
 * @brief Service implementation code of SpreadSheet.idl
 *
 */
#include "stdafx.h"
#include "SpreadSheetSVC_impl.h"

#include <vector>

#include "myExcel.h"
#include "MyObject.h"





mSpreadSheetSVC_impl::mSpreadSheetSVC_impl()
{

}


mSpreadSheetSVC_impl::~mSpreadSheetSVC_impl()
{

}


/*
 * セルの文字列を取得
 * l：行番号
 * c：列番号
 * sn：シート名
 * 戻り値：セルの文字列
 */
char* mSpreadSheetSVC_impl::get_string(const char* l, const char* c, const char* sn)
{
	std::string cans = myExcel::Obj->GetCellValueSingle(string2binary<int>(c,10),l, sn);
	//char *ans = const_cast <char *>(cans);
	std::cout << cans.c_str() << std::endl;
	
	return CORBA::string_dup(cans.c_str());
	

	
  
#ifndef WIN32
  #warning "Code missing in function <char* mSpreadSheetSVC_impl::get_string(const char* l, const char* c, const char* sn)>"
#endif
  //return 0;
}

/*
 * セルの値を設定
 * l：行番号
 * c：列番号
 * sn：シート名
 * v：設定する値
 */
void mSpreadSheetSVC_impl::set_value(const char* l, const char* c, const char* sn, CORBA::Float v)
{

	myExcel::Obj->SetCellValueSingle<float>(string2binary<int>(c,10),l, sn, v);
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_value(const char* l, const char* c, const char* sn, CORBA::Float v)>"
#endif
}

/*
 * 未実装
 */
StringList *mSpreadSheetSVC_impl::get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn)
{
  
#ifndef WIN32
  #warning "Code missing in function <SpreadSheet::StringList mSpreadSheetSVC_impl::get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn)>"
#endif
  return 0;
}

/*
 * 未実装
 */
void mSpreadSheetSVC_impl::set_value_range(const char* l, const char* c, const char* sn, FloatList v)
{
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_value_range(const char* l, const char* c, const char* sn, ValueList v)>"
#endif
}

/*
 * セルの文字列を設定
 * l：行番号
 * c：列番号
 * sn：シート名
 * v：設定する文字列
 */
void mSpreadSheetSVC_impl::set_string(const char* l, const char* c, const char* sn, const char* v)
{
	
	myExcel::Obj->SetCellValueSingle<System::String^>(string2binary<int>(c,10),l, sn, gcnew System::String(v));
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_string(const char* l, const char* c, const char* sn, const char* v)>"
#endif
}

/*
 * 未実装
 */
void mSpreadSheetSVC_impl::set_string_range(const char* l, const char* c, const char* sn, StringList v)
{
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_string_range(const char* l, const char* c, const char* sn, StringList v)>"
#endif
}







