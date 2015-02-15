// -*-C++-*-
/*!
 * @file  SpreadSheetSVC_impl.cpp
 * @brief Service implementation code of SpreadSheet.idl
 *
 */
#include "stdafx.h"
#include "SpreadSheetSVC_impl.h"

#include <vector>
#include <coil/stringutil.h>

#include "ExcelObject.h"
#include "SubFunction.h"







mSpreadSheetSVC_impl::mSpreadSheetSVC_impl()
{

}


mSpreadSheetSVC_impl::~mSpreadSheetSVC_impl()
{

}



char* mSpreadSheetSVC_impl::get_string(const char* l, const char* c, const char* sn)
{
	int c_tmp;
	coil::stringTo<int>(c_tmp,c);
	std::string cans = ExcelObject::Obj->getCellValueSingle(c_tmp,l, sn);
	//char *ans = const_cast <char *>(cans);
	std::cout << cans.c_str() << std::endl;
	
	return CORBA::string_dup(cans.c_str());
	

	
  
#ifndef WIN32
  #warning "Code missing in function <char* mSpreadSheetSVC_impl::get_string(const char* l, const char* c, const char* sn)>"
#endif
  //return 0;
}


void mSpreadSheetSVC_impl::set_value(const char* l, const char* c, const char* sn, CORBA::Float v)
{
	int c_tmp;
	coil::stringTo<int>(c_tmp,c);

	ExcelObject::Obj->setCellValueSingle<float>(c_tmp,l, sn, v);
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_value(const char* l, const char* c, const char* sn, CORBA::Float v)>"
#endif
}


StringList *mSpreadSheetSVC_impl::get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn)
{
  
#ifndef WIN32
  #warning "Code missing in function <SpreadSheet::StringList mSpreadSheetSVC_impl::get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn)>"
#endif
  return 0;
}


void mSpreadSheetSVC_impl::set_value_range(const char* l, const char* c, const char* sn, FloatList v)
{
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_value_range(const char* l, const char* c, const char* sn, ValueList v)>"
#endif
}


void mSpreadSheetSVC_impl::set_string(const char* l, const char* c, const char* sn, const char* v)
{
	int c_tmp;
	coil::stringTo<int>(c_tmp,c);

	ExcelObject::Obj->setCellValueSingle<System::String^>(c_tmp,l, sn, gcnew System::String(v));
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_string(const char* l, const char* c, const char* sn, const char* v)>"
#endif
}


void mSpreadSheetSVC_impl::set_string_range(const char* l, const char* c, const char* sn, StringList v)
{
  
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_string_range(const char* l, const char* c, const char* sn, StringList v)>"
#endif
}







