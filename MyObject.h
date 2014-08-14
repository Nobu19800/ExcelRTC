#ifndef MYOBJECT_H
#define MYOBJECT_H

#include <vector>
#include <string>
#include <sstream>


//セルの番号(アルファベット)を数値に変換する関数
int convertStrToVal(std::string m_str);


//文字を置換する関数
std::string Replace( std::string String1, std::string String2, std::string String3 );


//RTCのデータポートのツリーの各オブジェクトクラス
class TreeObject
{
public:
	TreeObject(std::string n)
	{
		name = n;
	};
	std::string name;
	std::vector<TreeObject*> to;
};

//System::Stringをstd::stringに変換する関数
std::string MarshalString ( System::String ^ s);

//文字列を特定の文字で分割する関数
std::vector<std::string> split(const std::string &str, const std::string &delim);

//文字列を数値に変換する関数
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