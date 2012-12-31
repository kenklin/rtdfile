#include <string>

#define SEP	"\t"

class Tuple
{
private:
	// Returns the position of the next argument.  Returns npos if no more.
	static size_t ParseArg(const std::string& str, size_t pos, std::string* out)
	{
		size_t next = std::string::npos;

		*out = "";
		if (str[pos] != '\0') {
			size_t sep = str.find(SEP, pos);
			if (sep == std::string::npos) {
				*out = str.substr(pos);
				next = str.length();	// Point to '\0'
			} else {
				*out = str.substr(pos, sep - pos);
				next = sep + 1;			// Point past sep
			}
		}

		return next;
	}

public:
	static std::string Create(const std::string& arg1, const std::string& arg2)
	{
		return arg1 + SEP + arg2;
	}

#if 0
	static std::string Create(const std::string& arg1, const std::string& arg2, const std::string& arg3)
	{
		return arg1 + SEP + arg2 + SEP + arg3;
	}
#endif

	static int Split(const std::string& args, std::string* arg1, std::string* arg2)
	{
		int		argc = 0;
		size_t	pos = 0;

		if ((pos = ParseArg(args, pos, arg1)) != std::string::npos) {
			argc++;
			if ((pos = ParseArg(args, pos, arg2)) != std::string::npos) {
				argc++;
			}
		}
		return argc;
	}

	static bool Get(const std::string& args, int n, std::string* arg)
	{
		int		argc = 0;
		size_t	pos = 0;

		while ((pos = ParseArg(args, pos, arg)) != std::string::npos) {
			if (argc == n) {
				break;		// found
			} else {
				*arg = "";
				argc++;
			}
		}

		return argc == n;
	}
};
