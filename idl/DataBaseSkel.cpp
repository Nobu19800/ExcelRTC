// -*- C++ -*-
/*!
 *
 * THIS FILE IS GENERATED AUTOMATICALLY!! DO NOT EDIT!!
 *
 * @file DataBaseSkel.cpp 
 * @brief DataBase server skeleton wrapper
 * @date Tue Aug 12 02:27:17 2014 
 *
 */

#include "DataBaseSkel.h"

#if defined ORB_IS_TAO
#  include "DataBaseC.cpp"
#  include "DataBaseS.cpp"
#elif defined ORB_IS_OMNIORB
#  include "DataBaseSK.cc"
#  include "DataBaseDynSK.cc"
#elif defined ORB_IS_MICO
#  include "DataBase.cc"
#  include "DataBase_skel.cc"
#elif defined ORB_IS_ORBIT2
#  include "DataBase-cpp-stubs.cc"
#  include "DataBase-cpp-skels.cc"
#elif defined ORB_IS_RTORB
#  include "OpenRTM-aist-decls.h"
#  include "DataBase-common.c"
#  include "DataBase-stubs.c"
#  include "DataBase-skels.c"
#  include "DataBase-skelimpl.c"
#else
#  error "NO ORB defined"
#endif

// end of DataBaseSkel.cpp
