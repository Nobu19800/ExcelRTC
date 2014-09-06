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




/*
 * Example implementational code for IDL interface SpreadSheet::mSpreadSheet
 */
mSpreadSheetSVC_impl::mSpreadSheetSVC_impl()
{
  // Please add extra constructor code here.
}


mSpreadSheetSVC_impl::~mSpreadSheetSVC_impl()
{
  // Please add extra destructor code here.
}


/*
 * Methods corresponding to IDL attributes and operations
 */
char* mSpreadSheetSVC_impl::get_string(const char* l, const char* c, const char* sn)
{
	std::string cans = myExcel::Obj->GetCellValueSingle(string2binary<int>(c,10),l, sn);
	//char *ans = const_cast <char *>(cans);
	std::cout << cans.c_str() << std::endl;
	
	return CORBA::string_dup(cans.c_str());
	

	
  // Please insert your code here and remove the following warning pragma
#ifndef WIN32
  #warning "Code missing in function <char* mSpreadSheetSVC_impl::get_string(const char* l, const char* c, const char* sn)>"
#endif
  //return 0;
}

void mSpreadSheetSVC_impl::set_value(const char* l, const char* c, const char* sn, CORBA::Float v)
{

	myExcel::Obj->SetCellValueSingle<float>(string2binary<int>(c,10),l, sn, v);
  // Please insert your code here and remove the following warning pragma
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_value(const char* l, const char* c, const char* sn, CORBA::Float v)>"
#endif
}

StringList *mSpreadSheetSVC_impl::get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn)
{
  // Please insert your code here and remove the following warning pragma
#ifndef WIN32
  #warning "Code missing in function <SpreadSheet::StringList mSpreadSheetSVC_impl::get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn)>"
#endif
  return 0;
}

void mSpreadSheetSVC_impl::set_value_range(const char* l, const char* c, const char* sn, FloatList v)
{
  // Please insert your code here and remove the following warning pragma
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_value_range(const char* l, const char* c, const char* sn, ValueList v)>"
#endif
}

void mSpreadSheetSVC_impl::set_string(const char* l, const char* c, const char* sn, const char* v)
{
	
	myExcel::Obj->SetCellValueSingle<System::String^>(string2binary<int>(c,10),l, sn, gcnew System::String(v));
  // Please insert your code here and remove the following warning pragma
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_string(const char* l, const char* c, const char* sn, const char* v)>"
#endif
}

void mSpreadSheetSVC_impl::set_string_range(const char* l, const char* c, const char* sn, StringList v)
{
  // Please insert your code here and remove the following warning pragma
#ifndef WIN32
  #warning "Code missing in function <void mSpreadSheetSVC_impl::set_string_range(const char* l, const char* c, const char* sn, StringList v)>"
#endif
}



// End of example implementational code



