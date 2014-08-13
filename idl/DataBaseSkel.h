// -*- C++ -*-
/*!
 *
 * THIS FILE IS GENERATED AUTOMATICALLY!! DO NOT EDIT!!
 *
 * @file DataBaseSkel.h 
 * @brief DataBase server skeleton header wrapper code
 * @date Tue Aug 12 02:27:17 2014 
 *
 */

#ifndef _DATABASESKEL_H
#define _DATABASESKEL_H



#include <rtm/config_rtc.h>
#undef PACKAGE_BUGREPORT
#undef PACKAGE_NAME
#undef PACKAGE_STRING
#undef PACKAGE_TARNAME
#undef PACKAGE_VERSION

#if   defined ORB_IS_TAO
#  include "DataBaseC.h"
#  include "DataBaseS.h"
#elif defined ORB_IS_OMNIORB
#  if defined(WIN32) || defined(_WIN32) || defined(__WIN32__) || defined(__NT__)
#    undef USE_stub_in_nt_dll
#  endif
#  include "DataBase.hh"
#elif defined ORB_IS_MICO
#  include "DataBase.h"
#elif defined ORB_IS_ORBIT2
#  include "/DataBase-cpp-stubs.h"
#  include "/DataBase-cpp-skels.h"
#elif defined ORB_IS_RTORB
#  include "DataBase.h"
#else
#  error "NO ORB defined"
#endif

#endif // _DATABASESKEL_H
