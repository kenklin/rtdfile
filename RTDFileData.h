#include <string>
#include <map>

class RTDFileData
{
public:
	RTDFileData() {}
	virtual ~RTDFileData() {}

public:
	std::string	LookupData(const std::string& Args);

private:
//	std::map<std::string, std::string> m_Data;
};
