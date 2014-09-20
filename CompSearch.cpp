// -*-C++-*-
/*!
 * @file  CompSearch.cpp
 * @brief
 *
 */

#include "stdafx.h"
#include "ExcelRTComponent.h"

#include <iostream>
#include <fstream>
#include "MyObject.h"

using namespace std;

/*
http://www17.atpages.jp/~aitech/wiki/?OpenRTM%2FControl%2Fconnect
http://www17.atpages.jp/~aitech/wiki/?OpenRTM%2FControl%2Flist
のソースコードを一部改変しました
*/



void portConnect(PortService_ptr p1, PortService_var p2)
{
	string subs_type = "flush";
	string period = "";

	p1->disconnect_all();
	p2->disconnect_all();

	string ConnectName = "_";
	ConnectName += p1->get_port_profile()->name;
	ConnectName += "_";
	ConnectName += p2->get_port_profile()->name;

	string ConnectIDName = "_id_";
	ConnectIDName += p1->get_port_profile()->name;
	ConnectIDName += "_";
	ConnectIDName += p2->get_port_profile()->name;


/*
http://www17.atpages.jp/~aitech/wiki/?OpenRTM%2FControl%2Fconnect
のソースコード70～78行目
*/

	RTC::ConnectorProfile prof;
	prof.connector_id = ConnectIDName.c_str();
	prof.name = CORBA::string_dup(ConnectName.c_str());
	prof.ports.length(2);
	prof.ports[0] = p1;
	prof.ports[1] = p2;


    CORBA_SeqUtil::push_back(prof.properties,NVUtil::newNV("dataport.interface_type", "corba_cdr"));
    CORBA_SeqUtil::push_back(prof.properties,NVUtil::newNV("dataport.dataflow_type","push"));
    if(subs_type=="flush")
        CORBA_SeqUtil::push_back(prof.properties,NVUtil::newNV("dataport.subscription_type","flush"));
    else if(subs_type=="new")
        CORBA_SeqUtil::push_back(prof.properties,NVUtil::newNV("dataport.subscription_type","new"));
    else if(subs_type=="periodic"){
        CORBA_SeqUtil::push_back(prof.properties,NVUtil::newNV("dataport.subscription_type","periodic"));
        CORBA_SeqUtil::push_back(prof.properties,NVUtil::newNV("dataport.push_rate",period.c_str()));
    }

/*ここまで*/

    RTC::ReturnCode_t ret;



    ret = p1->connect(prof);

}

/*
以下はhttp://www17.atpages.jp/~aitech/wiki/?OpenRTM%2FControl%2Flist
のソースコードを改変しています。
*/

/*引数にTreeObjectを追加*/
void ListRecursive(CosNaming::NamingContext_ptr context,vector<OtherPort> &rtclist,vector<string> &name, TreeObject *to){
    CosNaming::BindingList_var     bl;
    CosNaming::BindingIterator_var bi;
    CORBA::Boolean cont(true);
    int m_blLength=100;
    context->list(m_blLength, bl, bi);
 
    while (cont){
        CORBA::ULong len(bl->length());
        for (CORBA::ULong i = 0; i < len; ++i){
 
            if (bl[i].binding_type == CosNaming::ncontext){
                CosNaming::NamingContext_var next_context;
                next_context = CosNaming::NamingContext::_narrow(context->resolve(bl[i].binding_name));
                vector<string> namebuff=name;
                name.push_back(string(bl[i].binding_name[0].id));
                //name.push_back("/");
				TreeObject *to2 = new TreeObject(string(bl[i].binding_name[0].id));
				to->to.push_back(to2);
                ListRecursive(next_context,rtclist,name,to2); 
                name=namebuff;
            }
            else if (bl[i].binding_type == CosNaming::nobject){
                if(rtclist.size()>m_blLength)
                    break;
                vector<string> namebuff=name;
                name.push_back(string(bl[i].binding_name[0].id));
/*
データポートの情報を取得するためのコードを追加
*/
				if(string(bl[i].binding_name[0].kind) == "rtc")
				{
					
					RTC::CorbaConsumer<RTC::RTObject> rto;
					rto.setObject(context->resolve(bl[i].binding_name));
					try
					{
						RTC::PortServiceList_var tp = rto._ptr()->get_ports();
					
						TreeObject *to2 = new TreeObject(string(bl[i].binding_name[0].id));
						to->to.push_back(to2);

						for(int k=0;k < tp->length();k++)
						{
							vector<string> namebuff2=name;
							string tname = tp[(CORBA::ULong)k]->get_port_profile()->name;
							PortService_var p = tp[(CORBA::ULong)k];
							vector<string> pn = split(tname, ".");
							namebuff2.push_back(pn[1]);

							TreeObject *to3 = new TreeObject(string(pn[1]));
							to2->to.push_back(to3);

							rtclist.push_back(OtherPort(p, namebuff2));
						}
					}
					catch(...)
					{

					}

					
					

					//name=namebuff2;
				}
/*ここまで*/
				name=namebuff;
				
                //rtclist.push_back(name_buff);
            }
            else {
            }
        }
        if (CORBA::is_nil(bi)) {
            cont = 0;
        }
        else {
            bi->next_n(m_blLength, bl);
        }
    }
    return;
}


/*引数にTreeObjectとIPアドレスを追加*/
int rtc_get_rtclist(RTC::CorbaNaming &naming,vector<OtherPort> &rtclist, TreeObject *to, std::string IP_adress){
	
    
    CosNaming::NamingContext_ptr name_cxt = naming.getRootContext();
    rtclist.clear();
    vector<string> name;
	name.push_back(IP_adress);
    ListRecursive(name_cxt,rtclist,name,to);
    

    for(int i=0;i<(signed)rtclist.size();i++){
		std::string buff = "";
		for(int j=0;j < rtclist[i].buff.size();j++)
		{
			buff += rtclist[i].buff[j];
			buff += "/";
		}
		System::String^ a = gcnew System::String(buff.c_str());
		
        }
    
    return 0;
}

