#include <windows.h>
#include <mapi.h>
#include <sys/stat.h>
#include <crtdbg.h>
#include <string>
#include <vector>
#include <memory>

typedef std::vector<std::string> STRVEC;

// Simple wrapper around the mapi lib HINSTANCE + a couple calls.
// Makes dealing with the loading/unloading of the dll and the
// send call easier.
// Implemented at EOF.
class MapiLib {
public:
	MapiLib();
	~MapiLib();

	BOOL Load();
	ULONG Send(MapiMessage*);
	MapiRecipDesc* Resolve(const char* pszName) const;
	void FreeBuffer(void* pv) { m_pfnFreeBuff(pv); }

private:
	HINSTANCE m_hMapi;
	LPMAPISENDMAIL m_pfnSend;
	LPMAPIRESOLVENAME m_pfnResolve;
	LPMAPIFREEBUFFER  m_pfnFreeBuff;

	MapiLib(const MapiLib&); // no impl
	MapiLib& operator= (const MapiLib&); // no impl
};


// Here's the hassle this class tries to resolve:
// When we call MAPIResolveName, it allocates a MapiRecipDesc structure that we
// have to free with MAPIFreeBuffer. Could have a vector<MapiRecipDesc*> but
// that would only solve the problem of deleting the pointer. We'd still have to
// call MAPIFreeBuffer. So instead we'll have a vector of these which will do
// a MAPIFreeBuffer on the MapiRecipDesc* we're holding.
class RecipDescHolder {
public:
	RecipDescHolder(MapiLib& lib, MapiRecipDesc* p) : m_lib(lib), m_ptr(p) {}
	~RecipDescHolder() { if (m_ptr) m_lib.FreeBuffer(m_ptr); }
	MapiRecipDesc* Ptr() { return m_ptr; }

private:
	MapiLib&       m_lib;
	MapiRecipDesc* m_ptr;

	RecipDescHolder(const RecipDescHolder&); // no impl
	RecipDescHolder& operator= (const RecipDescHolder&); // no impl
};


// Parses the send data passed into the call.
// Implemented at EOF.
BOOL ParseSendData(
	const char* pszFile,
	STRVEC& vTo, 
	STRVEC& vCc, 
	STRVEC& vBcc, 
	std::string& sFrom, 
	std::string& sSubj, 
	std::string& sNote,
	STRVEC& vAttachments);


// Exported lib call
extern "C" _declspec(dllexport) char* VimSendMail(char* pszFile) {
	// Return string
	static char szRet[256];

	// Initial parse of incoming data
	std::string sFrom, sSubj, sNote;
	STRVEC vTo, vCc, vBcc, vAttachments;
	if (!ParseSendData((const char*)pszFile,
		vTo, vCc, vBcc, sFrom, sSubj, sNote, vAttachments))
		return "Failed to parse send data";

	// Load MAPI
	MapiLib lib;
	if (!lib.Load())
		return "Failed to load mapi32.dll";

	// Resolve sender
	RecipDescHolder recipFrom(lib, lib.Resolve(sFrom.c_str()));
	if (!recipFrom.Ptr())
		return "Failed to resolve sender address";

	recipFrom.Ptr()->ulRecipClass = MAPI_ORIG;

	// Resolve To, Cc, and Bcc
	int n, nTO = vTo.size(), nCC = vCc.size(), nBCC = vBcc.size();
	std::vector<RecipDescHolder*> vecRecips;
	MapiRecipDesc* pDesc;
	for (n = 0; n < nTO; n++) {
		pDesc = lib.Resolve(vTo[n].c_str());
		if (!pDesc) {
			sprintf(szRet, "Failed to resolve (to) address for %s",
				vTo[n].c_str());
			return szRet;
		}
		pDesc->ulRecipClass = MAPI_TO;
		vecRecips.push_back(new RecipDescHolder(lib, pDesc));
	}
	for (n = 0; n < nCC; n++) {
		pDesc = lib.Resolve(vCc[n].c_str());
		if (!pDesc) {
			sprintf(szRet, "Failed to resolve (cc) address for %s",
				vCc[n].c_str());
			return szRet;
		}
		pDesc->ulRecipClass = MAPI_CC;
		vecRecips.push_back(new RecipDescHolder(lib, pDesc));
	}
	for (n = 0; n < nBCC; n++) {
		pDesc = lib.Resolve(vBcc[n].c_str());
		if (!pDesc) {
			sprintf(szRet, "Failed to resolve (bcc) address for %s",
				vBcc[n].c_str());
			return szRet;
		}
		pDesc->ulRecipClass = MAPI_BCC;
		vecRecips.push_back(new RecipDescHolder(lib, pDesc));
	}

	// Put the MapiRecipDescs in a normal array
	MapiRecipDesc* pRecips = new MapiRecipDesc[nTO + nCC + nBCC];
	std::auto_ptr<MapiRecipDesc> spRecips(pRecips);
	memset(pRecips, 0, (nTO + nCC + nBCC) * sizeof(MapiRecipDesc));
	for (n = 0; n < (nTO + nCC + nBCC); n++) {
		pRecips[n].ulRecipClass = vecRecips[n]->Ptr()->ulRecipClass;
		pRecips[n].lpszName     = vecRecips[n]->Ptr()->lpszName;
		pRecips[n].lpszAddress  = vecRecips[n]->Ptr()->lpszAddress;
		pRecips[n].ulEIDSize    = vecRecips[n]->Ptr()->ulEIDSize;
		pRecips[n].lpEntryID    = vecRecips[n]->Ptr()->lpEntryID;
	}

	// Attachments?
	int nAttachments = vAttachments.size();
	MapiFileDesc* pAttachments = 0;
	std::auto_ptr<MapiFileDesc> spAttach(pAttachments);
	if (nAttachments) {
		pAttachments = new MapiFileDesc[nAttachments];
		memset(pAttachments, 0, sizeof(MapiFileDesc) * nAttachments);
		for (n = 0; n < nAttachments; n++) {
			pAttachments[n].lpszPathName = (char*)vAttachments[n].c_str();
			pAttachments[n].nPosition = (ULONG)-1;
		}
	}

	// Construct the message
	MapiMessage message;
	memset(&message, 0, sizeof(message));
	message.lpszSubject  = (char*)sSubj.c_str();
	message.lpszNoteText = (char*)sNote.c_str();
	message.lpOriginator = recipFrom.Ptr();
	message.lpRecips     = pRecips;
	message.nRecipCount  = nTO + nCC + nBCC;
	message.nFileCount   = nAttachments;
	message.lpFiles      = pAttachments;

	// And go
	ULONG rc = lib.Send(&message);
	if (0 == rc)
		strcpy(szRet, "Succeeded");
	else
		sprintf(szRet, "MAPI failure (%lu)", rc);

	return szRet;
}


extern "C" BOOL WINAPI DllMain(HINSTANCE hInst, DWORD dwReason, LPVOID) {
	switch(dwReason) {
	    case DLL_PROCESS_ATTACH:
			DisableThreadLibraryCalls(hInst); // couldn't care less...
			break;
//	    case DLL_THREAD_ATTACH:  // will no longer get this
//			break;
//		case DLL_THREAD_DETACH:  // ditto
//			break;
//	    case DLL_PROCESS_DETACH: // don't care
//			break;
	}
	return TRUE;
}


// Implement the MapiLib class
MapiLib::MapiLib()
	: m_hMapi(NULL),
	  m_pfnSend(NULL),
	  m_pfnResolve(NULL),
	  m_pfnFreeBuff(NULL)
{}


MapiLib::~MapiLib() {
	if (m_hMapi)
		FreeLibrary(m_hMapi);
	m_hMapi = NULL;
}


BOOL MapiLib::Load() {
	m_hMapi = LoadLibrary("MAPI32.DLL");
	if (!m_hMapi) {
		_RPT0(_CRT_WARN, "Failed to load mapi32.dll\n");
		return FALSE;
	}

	m_pfnSend = (LPMAPISENDMAIL)GetProcAddress(m_hMapi, "MAPISendMail");
	if (!m_pfnSend) {
		_RPT0(_CRT_WARN, "Failed to find MAPISendMail export in mapi32.dll\n");
		return FALSE;
	}

	m_pfnResolve = (LPMAPIRESOLVENAME)GetProcAddress(m_hMapi, "MAPIResolveName");
	if (!m_pfnResolve) {
		_RPT0(_CRT_WARN, "Failed to find MAPIResolveName export in mapi32.dll\n");
		return FALSE;
	}

	m_pfnFreeBuff = (LPMAPIFREEBUFFER)GetProcAddress(m_hMapi, "MAPIFreeBuffer");
	if (!m_pfnFreeBuff) {
		_RPT0(_CRT_WARN, "Failed to find MAPIFreeBuffer export in mapi32.dll\n");
		return FALSE;
	}

	return TRUE;
}


ULONG MapiLib::Send(MapiMessage* pMsg) {
	ULONG rc = m_pfnSend(0, 0, pMsg, 0, 0);
	_RPT1(_CRT_WARN, "rc from SendMail export: %lu\n", rc);
	return rc;
}


MapiRecipDesc* MapiLib::Resolve(const char* pszName) const {
	MapiRecipDesc *desc;
	ULONG rc = m_pfnResolve(0, 0, (char*)pszName, 0, 0, &desc);
	_RPT1(_CRT_WARN, "rc from ResolveName export: %lu\n", rc);
	return (SUCCESS_SUCCESS == rc) ? desc : NULL;
}


BOOL ParseSendData(
	const char* pszFile,
	STRVEC& vTo, 
	STRVEC& vCc, 
	STRVEC& vBcc, 
	std::string& sFrom, 
	std::string& sSubj, 
	std::string& sNote,
	STRVEC& vAttachments) {

	// How big is the file?
	struct _stat stats;
	if (0 != _stat(pszFile, &stats)) {
		_RPT1(_CRT_WARN, "Failed to get stats for file <%s>\n", pszFile);
		return FALSE;
	}

	// Alloc mem to hold the entire thing
	size_t nSize = stats.st_size;
	char* pMem = new char[nSize+1];
	if (!pMem) {
		_RPT1(_CRT_WARN, "Failed to alloc %d bytes to hold file\n", nSize);
		return FALSE;
	}

	std::auto_ptr<char> spMem(pMem); // this'll handle freeing it for us
	memset(pMem, 0, nSize+1);

	// Open for reading
	HANDLE hFile = CreateFile(pszFile, GENERIC_READ, FILE_SHARE_READ,
		NULL, OPEN_EXISTING, 0, NULL);
	if (INVALID_HANDLE_VALUE == hFile) {
		_RPT1(_CRT_WARN, "Failed to open file <%s>\n", pszFile);
		return FALSE;
	}

	// Read the entire thing
	DWORD dwRead;
	BOOL bOK = ReadFile(hFile, pMem, DWORD(nSize), &dwRead, NULL);
	CloseHandle(hFile);
	if (!bOK) {
		_RPT1(_CRT_WARN, "Failed to read file <%s>\n", pszFile);
		return FALSE;
	}

	// Contents of file now in pMem. It is nSize in length and file is closed

	// A brute force approach to parsing the header info.
	// A more elegant approach left for an exercise for the reader

	static const size_t knWORK = 1024;
	char szWork[knWORK];
	char* pt;
	size_t n = 0;
	size_t pos = 0; // current position in pMem
	static const char cCR = 0x0D;
	static const char cLF = 0x0A;

	// Get 'To:'
	bOK = FALSE;
	while ((pos < nSize) && (n < knWORK)) {
		szWork[n] = pMem[pos++];
		if (cLF == szWork[n]) {
			szWork[n] = '\0';
			if ((n > 0) && (cCR == szWork[n-1]))
				szWork[--n] = '\0';
			bOK = TRUE;
			break;
		}
		n++;
	}

	if ((!bOK) || (n < 5) || (strnicmp(szWork, "TO: ", 4))) {
		_RPT0(_CRT_WARN, "Failed to get TO: field\n");
		return FALSE;
	}

	pt = strtok(&szWork[4], ";");
	while ((pt) && (strchr(pt, '@'))) {
		_RPT1(_CRT_WARN, "to - %s\n", pt);
		vTo.push_back(std::string(pt));
		pt = strtok(NULL, ";");
	}

	// Must have at least one recipient
	if (vTo.size() < 1) {
		_RPT0(_CRT_WARN, "Failed to get TO: field\n");
		return FALSE;
	}


	// Get 'cc:'
	n = 0;
	bOK = FALSE;
	while ((pos < nSize) && (n < knWORK)) {
		szWork[n] = pMem[pos++];
		if (cLF == szWork[n]) {
			szWork[n] = '\0';
			if ((n > 0) && (cCR == szWork[n-1]))
				szWork[--n] = '\0';
			bOK = TRUE;
			break;
		}
		n++;
	}

	if ((!bOK) || (n < 4) || (strnicmp(szWork, "CC: ", 4))) {
		_RPT0(_CRT_WARN, "Failed to get CC: field\n");
		return FALSE;
	}

	if (n > 4) {
		pt = strtok(&szWork[4], ";");
		while ((pt) && (strchr(pt, '@'))) {
			_RPT1(_CRT_WARN, "cc - %s\n", pt);
			vCc.push_back(std::string(pt));
			pt = strtok(NULL, ";");
		}
	}


	// Get 'bcc:'
	n = 0;
	bOK = FALSE;
	while ((pos < nSize) && (n < knWORK)) {
		szWork[n] = pMem[pos++];
		if (cLF == szWork[n]) {
			szWork[n] = '\0';
			if ((n > 0) && (cCR == szWork[n-1]))
				szWork[--n] = '\0';
			bOK = TRUE;
			break;
		}
		n++;
	}

	if ((!bOK) || (n < 5) || (strnicmp(szWork, "BCC: ", 5))) {
		_RPT0(_CRT_WARN, "Failed to get BCC: field\n");
		return FALSE;
	}

	if (n > 5) {
		pt = strtok(&szWork[5], ";");
		while ((pt) && (strchr(pt, '@'))) {
			_RPT1(_CRT_WARN, "bcc - %s\n", pt);
			vBcc.push_back(std::string(pt));
			pt = strtok(NULL, ";");
		}
	}


	// Get 'From:'
	n = 0;
	bOK = FALSE;
	while ((pos < nSize) && (n < knWORK)) {
		szWork[n] = pMem[pos++];
		if (cLF == szWork[n]) {
			szWork[n] = '\0';
			if ((n > 0) && (cCR == szWork[n-1]))
				szWork[--n] = '\0';
			bOK = TRUE;
			break;
		}
		n++;
	}

	if ((bOK) && (n > 6) && (!strnicmp(szWork, "FROM: ", 6))) {
		sFrom = &szWork[6];
	}
	else {
		_RPT0(_CRT_WARN, "Failed to get FROM: field\n");
		return FALSE;
	}


	// Get 'Subject:'
	n = 0;
	bOK = FALSE;
	while ((pos < nSize) && (n < knWORK)) {
		szWork[n] = pMem[pos++];
		if (cLF == szWork[n]) {
			szWork[n] = '\0';
			if ((n > 0) && (cCR == szWork[n-1]))
				szWork[--n] = '\0';
			bOK = TRUE;
			break;
		}
		n++;
	}

	if ((!bOK) || (n < 9) || (strnicmp(szWork, "SUBJECT: ", 9))) {
		_RPT0(_CRT_WARN, "Failed to get SUBJECT: field\n");
		return FALSE;
	}

	if (n > 9)
		sSubj = &szWork[9];

	// Whatever remains is the message text. The "AttachFile[...]"
	// markers, if any, are included in the message text.
	if (pos < nSize) {
		sNote = &pMem[pos];

		// Any attachments?
		pt = strstr(&pMem[pos], "AttachFile[");
		while (pt) {
			pt = strchr(pt, '[');
			pt++;
			n = strlen(pt);
			if ((n) && (strchr(pt, ']'))) {
				memset(szWork, 0, sizeof(szWork));
				n = 0;
				while (n < _MAX_PATH) {
					szWork[n] = pt[n];
					if (']' == szWork[n]) {
						szWork[n] = '\0';
						bOK = TRUE;
						break;
					}
					n++;
				}
				if (bOK && n)
					vAttachments.push_back(szWork);
			}
			pt = strstr(pt, "AttachFile[");
		}
	}

	return TRUE;
}

