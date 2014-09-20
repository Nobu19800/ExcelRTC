// -*-C++-*-
/*!
 * @file  SpreadSheetSVC_impl.h
 * @brief Service implementation header of SpreadSheet.idl
 *
 */

#include "SpreadSheetSkel.h"

#ifndef SPREADSHEETSVC_IMPL_H
#define SPREADSHEETSVC_IMPL_H
 
/*!
 * @class mSpreadSheetSVC_impl
 */
class mSpreadSheetSVC_impl
 : public virtual POA_SpreadSheet::mSpreadSheet,
   public virtual PortableServer::RefCountServantBase
{
 private:
   

 public:
  /*!
   * @brief standard constructor
   */
   mSpreadSheetSVC_impl();
  /*!
   * @brief destructor
   */
   virtual ~mSpreadSheetSVC_impl();

   // attributes and operations
   /**
   *@brief セルの文字列を取得
   * @param l 行番号
   * @param c 列番号
   * @param sn シート名
   * @return セルの文字列
   */
   char* get_string(const char* l, const char* c, const char* sn);
   /**
   *@brief セルの値を設定
   * @param l 行番号
   * @param c 列番号
   * @param sn シート名
   * @param v 設定する値
   */
   void set_value(const char* l, const char* c, const char* sn, CORBA::Float v);
   /**
   *@brief 未実装
   */
   StringList* get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn);
   /**
   *@brief 未実装
   */
   void set_value_range(const char* l, const char* c, const char* sn, const FloatList v);
   /**
   *@brief セルの文字列を設定
   * @param l 行番号
   * @param c 列番号
   * @param sn シート名
   * @param v 設定する文字列
   */
   void set_string(const char* l, const char* c, const char* sn, const char* v);
   /**
   *@brief 未実装
   */
   void set_string_range(const char* l, const char* c, const char* sn, const StringList v);
   

};



#endif // SPREADSHEETSVC_IMPL_H


