#define WIN32_LEAN_AND_MEAN
#include <vector>
#include <utility>
#include <tuple>
#include <algorithm>
#include <Windows.h>

LPWSTR GetExtPart(LPWSTR str)
{
	const size_t length = wcslen(str);
	for (size_t n = length; n > 0;)
	{
		--n;
		if (str[n] == '.')
		{
			return str + n;
		}
		else if (str[n] == '\\' || str[n] == '/')
		{
			break;
		}
	}
	return str + length;
}

void RemoveTrailWhiteSpace(std::wstring& s)
{
	size_t n = s.find_last_not_of(L" \t\r\n");
	if (n != std::wstring::npos && n < s.size())
	{
		s.resize(n + 1);
	}
	else
	{
		s.clear();
	}
}

void RemoveFirstWhiteSpace(std::wstring& s)
{
	size_t n = s.find_first_not_of(L" \t\r\n");
	if (n != 0 && n != std::wstring::npos)
	{
		s.erase(0, n);
	}
}

void RemoveFirstArgument(std::wstring& result)
{
	RemoveFirstWhiteSpace(result);
	if (result.empty())
	{
		return;
	}

	if (result[0] == '\"')
	{
		bool found = false;
		const size_t length = result.size();
		size_t at = 1;
		for (; at < length; ++at)
		{
			const wchar_t c = result[at];
			if (c == '\\')
			{
				++at;
			}
			else if (c == '\"')
			{
				found = true;
				++at;
				break;
			}
		}

		if (found)
		{
			if (at < length)
			{
				result.erase(0, at);
				RemoveFirstWhiteSpace(result);
			}
		}
	}
	else
	{
		size_t at = result.find_first_of(L" \t\r\n");
		if (at != std::wstring::npos)
		{
			result.erase(0, at);
		}
		else
		{
			result.clear();
		}
	}
}

void BuildCommandLine(std::wstring& result, LPCWSTR appName, LPCWSTR inCmdLine)
{
	const size_t appNameLength = wcslen(appName);
	const size_t inCmdLineLength = wcslen(inCmdLine);

	const size_t reserveSize = appNameLength + inCmdLineLength + 4;

	result.clear();
	result.reserve(reserveSize);
	result.assign(inCmdLine);
	RemoveFirstArgument(result);
	RemoveTrailWhiteSpace(result);

	const bool hasArgument = !result.empty();

	result.insert(0, L"\"", 1);
	result.insert(1, appName, appNameLength);
	result.insert(appNameLength + 1, L"\"", 1);
	if (hasArgument)
	{
		result.insert(appNameLength + 2, L" ", 1);
	}

	const size_t n = result.size();
	for (size_t i = 0; i < n; ++i)
	{
		wchar_t c = result[i];
		if (c == '\t' || c == '\r' || c == '\n')
		{
			result[i] = ' ';
		}
	}

	result.append(L"\0");
}

enum class OverrideMethod
{
	PreserveExisting,
	InsertBefore,
	InsertAfter,
	Overwrite,
};

typedef std::pair<std::wstring, std::wstring> EnvironmentItem;
typedef std::tuple<
	std::wstring, std::wstring, OverrideMethod> EnvironmentOverrideItem;

LPCWSTR ParseKeyValue(
	LPCWSTR p, std::vector<EnvironmentItem>& destination)
{
	LPCWSTR keyEnd = p;
	LPCWSTR valueBegin = p;
	LPCWSTR valueEnd = p;
	for (; *valueEnd; ++valueEnd)
	{
		if (*valueEnd == '=')
		{
			++valueBegin;
			break;
		}
		else
		{
			++valueBegin;
			++keyEnd;
		}
	}
	while (*valueEnd)
	{
		++valueEnd;
	}

	destination.emplace_back(EnvironmentItem(
		std::wstring(p, keyEnd),
		std::wstring(valueBegin, valueEnd)));

	return valueEnd + 1;
}

void ParseEnvironmentBlock(LPCWSTR environments,
	std::vector<EnvironmentItem>& destination)
{
	if (!environments)
	{
		return;
	}

	LPCWSTR p = environments;
	while (*p)
	{
		p = ParseKeyValue(p, destination);

#if defined(DEBUG_TRACE)
		wprintf(L"K:\"%s\", V:\"%s\"\n",
			result.first.c_str(),
			result.second.c_str());
#endif
	}
}

void ParseEnvironmentOverrideItem(
	LPCWSTR iniPath, LPCWSTR sectionName,
	std::vector<EnvironmentOverrideItem>& destination)
{
	enum { MaxLength = 32768 };
	WCHAR buffer[MaxLength] = {};
	::GetPrivateProfileStringW(
		sectionName, L"Key", NULL, buffer, _countof(buffer), iniPath);
	if (buffer[0] == '\0')
	{
		return;
	}

	std::wstring key(buffer);

	memset(buffer, 0, sizeof(buffer));
	::GetPrivateProfileStringW(
		sectionName, L"Value", NULL, buffer, _countof(buffer), iniPath);

	std::wstring value(buffer);

	memset(buffer, 0, sizeof(buffer));
	::GetPrivateProfileStringW(
		sectionName, L"Method", NULL, buffer, _countof(buffer), iniPath);

	LPWSTR endOfNumber = NULL;
	const int number = wcstol(buffer, &endOfNumber, 10);

	OverrideMethod method = OverrideMethod::PreserveExisting;

	if (endOfNumber == buffer + wcslen(buffer))
	{
		switch (static_cast<OverrideMethod>(number))
		{
		case OverrideMethod::PreserveExisting:
		case OverrideMethod::InsertBefore:
		case OverrideMethod::InsertAfter:
		case OverrideMethod::Overwrite:
			method = static_cast<OverrideMethod>(number);
			break;
		}
	}
	else if (_wcsicmp(buffer, L"InsertBefore") == 0)
	{
		method = OverrideMethod::InsertBefore;
	}
	else if (_wcsicmp(buffer, L"InsertAfter") == 0)
	{
		method = OverrideMethod::InsertAfter;
	}
	else if (_wcsicmp(buffer, L"Overwrite") == 0)
	{
		method = OverrideMethod::Overwrite;
	}

	destination.emplace_back(
		EnvironmentOverrideItem(
		std::move(key), std::move(value), std::move(method)));
}

void ParseEnvironmentOverrideList(
	LPCWSTR iniPath, LPCWSTR sectionNames,
	std::vector<EnvironmentOverrideItem>& destination)
{
	static const WCHAR prefix[] = L"EnvironmentVariable";
	static const int prefixLength = _countof(prefix) - 1;

	LPCWSTR section = sectionNames;
	while (*section)
	{
		if (_wcsnicmp(section, prefix, prefixLength) == 0)
		{
			ParseEnvironmentOverrideItem(iniPath, section, destination);
		}

		size_t n = wcslen(section);
		section += n + 1;
	}
}

void InitOption(
	std::wstring& targetExecutablePath,
	std::vector<EnvironmentOverrideItem>& environmentOverride)
{
	enum { MaxLength = 32768 };
	WCHAR iniPath[MaxLength] = {};
	{
		WCHAR moduleName[MaxLength] = {};
		::GetModuleFileNameW(
			::GetModuleHandleW(NULL), moduleName, _countof(moduleName));
		::GetFullPathNameW(moduleName, _countof(iniPath), iniPath, NULL);
	}

	LPWSTR extPart = GetExtPart(iniPath);
	size_t extPartCapacity = iniPath + _countof(iniPath) - extPart;
	wcsncpy_s(extPart, extPartCapacity, L".ini", _TRUNCATE);

	{
		WCHAR targetAppPath[MaxLength] = {};
		::GetPrivateProfileStringW(
			L"Executable", L"Path", L"", targetAppPath,
			_countof(targetAppPath), iniPath);
		targetExecutablePath = targetAppPath;
	}

	{
		WCHAR sectionNames[MaxLength] = {};
		::GetPrivateProfileSectionNamesW(
			sectionNames, _countof(sectionNames), iniPath);
		ParseEnvironmentOverrideList(
			iniPath, sectionNames, environmentOverride);
	}
}

void OverrideEnvironments(
	const std::vector<EnvironmentOverrideItem> environmentOverride,
	std::vector<EnvironmentItem>& targetEnvironment)
{
	for (auto& itOverride : environmentOverride)
	{
		auto& key = std::get<0>(itOverride);
		auto& value = std::get<1>(itOverride);
		auto& type = std::get<2>(itOverride);

		auto where = std::find_if(
			targetEnvironment.begin(), targetEnvironment.end(),
			[&itOverride, &key](decltype(*cbegin(targetEnvironment))& x)
		{
			return _wcsicmp(key.c_str(), x.first.c_str()) == 0;
		});

		if (where == targetEnvironment.end())
		{
			targetEnvironment.emplace_back(std::make_pair(key, value));
		}
		else if (type == OverrideMethod::InsertBefore)
		{
			(*where).second.insert(0, value);
		}
		else if (type == OverrideMethod::InsertAfter)
		{
			(*where).second.append(value);
		}
		else if (type == OverrideMethod::Overwrite)
		{
			(*where).second = value;
		}
	}
}

void BuildEnvironmentBlock(
	const std::vector<EnvironmentItem>& items,
	std::wstring& destination)
{
	if (items.empty())
	{
		return;
	}

	for (auto& i : items)
	{
		destination.append(i.first);
		destination.push_back(L'=');
		destination.append(i.second);
		destination.push_back(L'\0');
	}
	destination.push_back(L'\0');
}

int wmain()
{
	std::wstring targetExecutablePath;
	std::vector<EnvironmentOverrideItem> environmentOverride;
	InitOption(targetExecutablePath, environmentOverride);

	std::vector<EnvironmentItem> sourceEnvironment;
	LPWCH penv = ::GetEnvironmentStringsW();
	if (penv)
	{
		ParseEnvironmentBlock(penv, sourceEnvironment);
		::FreeEnvironmentStringsW(penv);
	}

	std::wstring commandLine;
	BuildCommandLine(
		commandLine, targetExecutablePath.c_str(), ::GetCommandLineW());

	std::wstring environmentOverrideBlock;
	if (!environmentOverride.empty())
	{
		OverrideEnvironments(environmentOverride, sourceEnvironment);
		BuildEnvironmentBlock(sourceEnvironment, environmentOverrideBlock);
	}

	LPVOID lpEnvironment =
		environmentOverrideBlock.empty() ? NULL : &environmentOverrideBlock[0];
	STARTUPINFOW StartupInfo = {};
	PROCESS_INFORMATION ProcessInformation = {};

	StartupInfo.cb = sizeof(StartupInfo);

#if defined(DEBUG_TRACE)
	wprintf(L"inCmdLine : \"%s\"\n", ::GetCommandLineW());
	wprintf(L"targetExecutablePath : \"%s\"\n", targetExecutablePath.c_str());
	wprintf(L"commandLine : \"%s\"\n", commandLine.c_str());
#endif

	BOOL created = ::CreateProcessW(
		NULL,                   // LPCWSTR lpApplicationName
		&commandLine[0],        // LPWSTR lpCommandLine
		NULL,                   // LPSECURITY_ATTRIBUTES lpProcessAttributes
		NULL,                   // LPSECURITY_ATTRIBUTES lpThreadAttributes
		TRUE,                   // BOOL bInheritHandles
		CREATE_UNICODE_ENVIRONMENT, // DWORD dwCreationFlags
		lpEnvironment,          // LPVOID lpEnvironment
		NULL,                   // LPCTSTR lpCurrentDirectory
		&StartupInfo,           // LPSTARTUPINFOW lpStartupInfo
		&ProcessInformation);   // LPPROCESS_INFORMATION lpProcessInformation

	int result = 0;
	if (created != FALSE)
	{
		if (ProcessInformation.hThread != NULL)
		{
			::CloseHandle(ProcessInformation.hThread);
		}

		::WaitForSingleObject(ProcessInformation.hProcess, INFINITE);
		DWORD ExitCode = 0;
		if (::GetExitCodeProcess(ProcessInformation.hProcess, &ExitCode))
		{
			result = ExitCode;
		}

		if (ProcessInformation.hProcess != NULL)
		{
			::CloseHandle(ProcessInformation.hProcess);
		}
	}
	else
	{
		result = ::GetLastError();
	}

	return result;
}

