// -*-C++-*-
/*!
 * @file  MyObject.h
 *
 */
#ifndef MYOBJECT_H
#define MYOBJECT_H

#include <vector>
#include <string>
#include <sstream>



/**
*@brief セルの番号(アルファベット)を数値に変換する関数
* @param m_str
* @return 
*/
int convertStrToVal(std::string m_str);



/**
*@brief 文字を置換する関数
* @param String1
* @param String2
* @param String3
* @return 
*/
std::string Replace( std::string String1, std::string String2, std::string String3 );



/**
* @class TreeObject
*@brief RTCのデータポートのツリーの各オブジェクトクラス
*/
class TreeObject
{
public:
	/**
	*@brief コンストラクタ
	* @param n
	*/
	TreeObject(std::string n)
	{
		name = n;
	};
	std::string name; /**<　@brief  */
	std::vector<TreeObject*> to; /**<　@brief  */
};


/**
*@brief System::Stringをstd::stringに変換する関数
* @param s
* @return 
*/
std::string MarshalString ( System::String ^ s);


/**
*@brief 文字列を特定の文字で分割する関数
* @param str
* @param delim
* @return 
*/
std::vector<std::string> split(const std::string &str, const std::string &delim);


/**
*@brief 文字列を数値に変換する関数
* @param text
* @param base
* @return 
*/
template<typename T>
T string2binary(const std::string& text, int base) {
    
    std::istringstream is(text);
    T value;
    switch (base) {
    case 8:
        is >> std::oct >> value;
        break;
    case 10:
        is >> value;
        break;
    case 16:
        is >> std::hex >> value;
        break;
    default:
        ;
    }

    return value;
}

#endif