// -*-C++-*-
/*!
 * @file  DataBaseSVC_impl.h
 * @brief Service implementation header of DataBase.idl
 *
 */

#include "DataBaseSkel.h"

#ifndef DATABASESVC_IMPL_H
#define DATABASESVC_IMPL_H
 
/*!
 * @class mDataBaseSVC_impl
 * Example class implementing IDL interface DataBase::mDataBase
 */
class mDataBaseSVC_impl
 : public virtual POA_DataBase::mDataBase,
   public virtual PortableServer::RefCountServantBase
{
 private:
   // Make sure all instances are built on the heap by making the
   // destructor non-public
   //virtual ~mDataBaseSVC_impl();

 public:
  /*!
   * @brief standard constructor
   */
   mDataBaseSVC_impl();
  /*!
   * @brief destructor
   */
   virtual ~mDataBaseSVC_impl();

   // attributes and operations
   char* get_string(const char* l, const char* c, const char* sn);
   void set_value(const char* l, const char* c, const char* sn, CORBA::Float v);
   StringList *get_string_range(const char* l1, const char* c1, const char* l2, const char* c2, const char* sn);
   void set_value_range(const char* l, const char* c, const char* sn, ValueList v);
   void set_string(const char* l, const char* c, const char* sn, const char* v);
   void set_string_range(const char* l, const char* c, const char* sn, StringList v);

};



#endif // DATABASESVC_IMPL_H


