// LocalStgDump.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "stdafx.h"
#include <WTypes.h>
#include <iostream>
#include <sstream>
#include <atltime.h>
#include <memory>

class TraverseInfo 
{
public:
	TraverseInfo();
	int m_nStorages;
	int m_nStreams;
	int m_nLockbytes;
	int m_nProperties;
	ULARGE_INTEGER m_cbTotalStreamSize;
	HRESULT m_hResult;
};

TraverseInfo::TraverseInfo()
	: m_nStorages(0), m_nStreams(0), 
	  m_nLockbytes(0), m_nProperties(0), m_hResult(S_OK)
{
	m_cbTotalStreamSize.HighPart = 0;   // Currently ignored; our total stream sizes are not greater than 4G
	m_cbTotalStreamSize.LowPart = 0;
}

template<class T> class ObjWrapper
{
public:
	ObjWrapper();
	~ObjWrapper();
	void Set(T* pObj);
private:
	T* m_pObj;
};

template<class T>
ObjWrapper<T>::ObjWrapper()
	: m_pObj(nullptr)
{}

template<class T>
ObjWrapper<T>::~ObjWrapper()
{
	if (m_pObj)
	{
		m_pObj->Release();
		m_pObj = nullptr;
	}
}

template<class T>
void ObjWrapper<T>::Set(T* pObj)
{
	m_pObj = pObj;
}


namespace LocStg {

	
	void ListTopLevel(const std::wstring& fileToLoad);
	void DisplayTotalStorageCount(const std::wstring& fileToLoad);
	void Traverse(IStorage* pStorage, TraverseInfo& ti);
	void Generate(const std::wstring& fileToWrite, int nParts);

	HRESULT OpenRootStorage(const std::wstring& strName, IStorage** o_ppStorage, bool bSilent=false);
	HRESULT GetStorageElementCount(IStorage* pStorage, int& elemCnt);
	HRESULT FindLocalDocsStorage(IStorage *pParent, IStorage **o_ppLocalDocsStg);
	HRESULT ListImmediateChildren(IStorage *pParent, int nChildren, int indentLevel);
	HRESULT CreateRootStorage(const std::wstring& strName, bool bOverwrite, IStorage** o_ppStorage);
	HRESULT CreatePartStorage(IStorage* pStorage, const std::wstring& partName);
	HRESULT CreatePartSubStorage(IStorage* pStorage, const std::wstring& stgName);

	void _Usage(const wchar_t* thisProg);
	void _WaitForKey();
	void _WriteIndent(int nLevels);
	void _WriteStorageType(DWORD type);

	constexpr wchar_t* sLocalDocs = L"LocalDocs";
	constexpr int nPartStorages = 15;
	constexpr int nPartWriteBlocks = 30;
	constexpr int nPartWriteBlockSize = 512;
}



int wmain(int argc, wchar_t* argv[])
{
	if (argc < 2)
	{
		LocStg::_Usage(argv[0]);
		return 0;
	}
	// Pause here to allow attaching debugger
	//LocStg::_WaitForKey();

	std::wstring fileName = argv[1];

	int opcode = 0;
	if (argc > 2)
	{
		opcode = _wtoi(argv[2]);
	}

	int partCount = 0;
	if (argc > 3)
	{
		partCount = _wtoi(argv[3]);
	}

	switch (opcode)
	{
		case 1: LocStg::ListTopLevel(fileName); break;
		case 2: LocStg::DisplayTotalStorageCount(fileName); break;
		case 3: LocStg::Generate(fileName, partCount); break;
		default: LocStg::_Usage(argv[0]);  break;
	}
}

void LocStg::ListTopLevel(const std::wstring& fileToLoad)
{
	std::wcout << L"Attempting to read " << fileToLoad.c_str() << std::endl;

	CTime startTime = CTime::GetCurrentTime();
	IStorage* pRoot = nullptr;
	HRESULT hr = LocStg::OpenRootStorage(fileToLoad, &pRoot);
	CTime endTime = CTime::GetCurrentTime();
	if (!SUCCEEDED(hr) || !pRoot)
	{
		std::wcout << L"Open root storage FAILED with error code " << hr << std::endl;
		return;
	}
	ObjWrapper<IStorage> w;
	w.Set(pRoot);

	std::wcout << L"Open root storage SUCCEEDED" << std::endl;
	CTimeSpan elapsedTime = endTime - startTime;
	std::wcout << L"Operation took " << elapsedTime.GetTotalSeconds() << L" seconds." << std::endl;

	int rootElemCnt = 0;
	hr = LocStg::GetStorageElementCount(pRoot, rootElemCnt);
	if (!SUCCEEDED(hr))
	{
		std::wcout << L"GetStorageElementCount failed with error code " << hr << std::endl;
	}

	STATSTG rootStatStg;
	pRoot->Stat(&rootStatStg, 0);
	int indentLevel = 0;
	_WriteStorageType(rootStatStg.type);
	std::wcout << L" \"" << rootStatStg.pwcsName << L"\" has " << rootElemCnt << L" elements" << std::endl;

	hr = LocStg::ListImmediateChildren(pRoot, rootElemCnt, indentLevel+1);

	// Pause before releasing storage to allow examination in ProcessExplorer etc.
	//_WaitForKey();
}

void LocStg::DisplayTotalStorageCount(const std::wstring& fileToLoad)
{
	IStorage* pRoot = nullptr;
	HRESULT hr = LocStg::OpenRootStorage(fileToLoad, &pRoot, true);
	if (!SUCCEEDED(hr) || !pRoot)
	{
		std::wcout << L"Open root storage FAILED with error code " << hr << std::endl;
		return;
	}

	TraverseInfo ti;
	Traverse(pRoot, ti);

	pRoot->Release();
	pRoot = nullptr;

	std::wcout << std::endl << std::endl << L"Traverse status for Storage " << fileToLoad.c_str();
	if (SUCCEEDED(ti.m_hResult))
	{
		std::wcout << L"  SUCCESS" << std::endl;
	}
	else 
	{
		std::wcout << L"  FAILED - error code = " << std::hex << ti.m_hResult << std::endl;
	}

	std::wcout << L"Total counts"  << std::endl;
	std::wcout << L"Storages:     " << ti.m_nStorages << std::endl;
	std::wcout << L"Streams:      " << ti.m_nStreams << std::endl;
	std::wcout << L"Stream bytes: " << ti.m_cbTotalStreamSize.LowPart << std::endl;
	std::wcout << L"Lockbytes:    " << ti.m_nLockbytes << std::endl;
	std::wcout << L"Properties:   " << ti.m_nProperties << std::endl;
}

void LocStg::Generate(const std::wstring& fileToWrite, int nParts)
{
	// Generates a data file containing:
	// fileToWrite STORAGE
	//     LocalDocs STORAGE
	//         ld<n>    'nParts' storages numbered ld0 ... ln<nParts-1>
	//         Each ld<n> storage corresponds to a "part"
	//            A part consists of nPartStorages
	//            Each of the nPartStorages has a stream of size nPartWriteBlocks * nPartWriteBlockSize
	// So the total number of storages = nParts * nPartStorages + 2 (for root, "LocalDocs)
	// Total stream bytes = nParts * nPartStorages * nPartWriteBlocks * nPartWriteBlockSize

	// Create the root storage
	IStorage* pRoot = nullptr;
	HRESULT hr = LocStg::CreateRootStorage(fileToWrite, true, &pRoot);
	if (!SUCCEEDED(hr) || !pRoot)
	{
		std::wcout << L"Could not create root storage \"" << fileToWrite.c_str() << "\" for writing" << std::endl;
		return;
	}
	ObjWrapper<IStorage> w;
	w.Set(pRoot);

	// Create "LocalDocs" storage below the root
	DWORD grfMode = STGM_WRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_DIRECT;
	IStorage* pLocalDocsStg = nullptr;
	pRoot->CreateStorage(LocStg::sLocalDocs, grfMode, 0, 0, &pLocalDocsStg);
	if (!SUCCEEDED(hr) || !pRoot)
	{
		std::wcout << L"Could not create storage \"" << LocStg::sLocalDocs << "\" for writing" << std::endl;
		return;
	}
	ObjWrapper<IStorage> w2;
	w2.Set(pLocalDocsStg);

	constexpr wchar_t* partPrefix = L"PART_";

	// Create 'nParts' parts below that
	for (int ii = 0; ii < nParts; ++ii)
	{
		std::wostringstream convHelper;
		convHelper << partPrefix << ii;
		std::wstring partName = convHelper.str();

		LocStg::CreatePartStorage(pLocalDocsStg, partName);
	}
}

HRESULT LocStg::CreatePartStorage(IStorage* pStorage, const std::wstring& partName)
{
	DWORD grfMode = STGM_WRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_DIRECT;
	IStorage* pPartStg = nullptr;
	HRESULT hr = pStorage->CreateStorage(partName.c_str(), grfMode, 0, 0, &pPartStg);
	if (!SUCCEEDED(hr) || !pPartStg)
	{
		std::wcout << L"Could not create storage \"" << partName.c_str() << "\" for writing" << std::endl;
		return hr;
	}
	ObjWrapper<IStorage> w;
	w.Set(pPartStg);

	// Generate the part's sub-storages
	constexpr wchar_t* partPrefix = L"PStg_";
	for (int ii = 0; ii < LocStg::nPartStorages; ++ii)
	{
		std::wostringstream convHelper;
		convHelper << partPrefix << ii;
		std::wstring stgName = convHelper.str();
		hr = LocStg::CreatePartSubStorage(pPartStg, stgName);
		if (!SUCCEEDED(hr)) break;
	}
	return hr;
}

static unsigned char* _PartData(size_t cb)
{
	static size_t currentSize = 0;
	static unsigned char* pData = nullptr;

	if (cb > currentSize)
	{
		if (pData) delete pData;
		pData = new unsigned char[cb];
		currentSize = cb;
		unsigned char* pStart = pData;

		for (size_t ii = 0; ii < currentSize; ++ii)
		{
			*pStart++ = ii % 256;
		}
	}
	return pData;
}

HRESULT LocStg::CreatePartSubStorage(IStorage* pStorage, const std::wstring& stgName)
{
	DWORD grfMode = STGM_WRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_DIRECT;
	IStorage* pSubStg = nullptr;
	HRESULT hr = pStorage->CreateStorage(stgName.c_str(), grfMode, 0, 0, &pSubStg);
	if (!SUCCEEDED(hr) || !pSubStg)
	{
		std::wcout << L"Could not create storage \"" << stgName.c_str() << "\" for writing" << std::endl;
		return hr;
	}
	ObjWrapper<IStorage> w;
	w.Set(pSubStg);

	std::wstring streamName = stgName + L"_Stream";
	// Create the stream inside the sub-storage
	IStream* pStream = nullptr;
	hr = pSubStg->CreateStream(streamName.c_str(), grfMode, 0, 0, &pStream);
	if (!SUCCEEDED(hr) || !pStream)
	{
		std::wcout << L"Could not create stream \"" << streamName.c_str() << "\" for writing" << std::endl;
		return hr;
	}
	ObjWrapper<IStream> sw;
	sw.Set(pStream);

	void const *pv = _PartData(LocStg::nPartWriteBlockSize);

	for (int ii = 0; ii < LocStg::nPartWriteBlocks; ++ii)
	{
		ULONG cbWritten = 0;
		hr = pStream->Write(pv, LocStg::nPartWriteBlockSize, &cbWritten);
		if (!SUCCEEDED(hr) || cbWritten < LocStg::nPartWriteBlockSize) break;
	}
	return hr;
}


// Traverse starting at pStorage, collecting storage and stream counts
void LocStg::Traverse(IStorage* pStorage, TraverseInfo& ti)
{
	// Visit this node (must be a storage)
	++(ti.m_nStorages);

	// Get an enumerator to this storage's immediate children
	IEnumSTATSTG *pEnum = nullptr;
	HRESULT hr = pStorage->EnumElements(0, nullptr, 0, &pEnum);
	if (!SUCCEEDED(hr) || !pEnum)
	{
		ti.m_hResult = hr;
		return;
	}
	ObjWrapper<IEnumSTATSTG> w;
	w.Set(pEnum);

	constexpr ULONG BlockSize = 100;
	STATSTG arrayStg[BlockSize];
	ULONG actualCount = 0;
	while (true)
	{
		// Get the next block of children to work on
		hr = pEnum->Next(BlockSize, arrayStg, &actualCount);
		if (!(SUCCEEDED(hr)))
		{
			ti.m_hResult = hr;
			return;
		}

		// Traverse that block of children
		STATSTG *pStatStg = &arrayStg[0];
		for (ULONG ii = 0; ii < actualCount; ++ii)
		{
			if (STGTY_STORAGE == pStatStg->type)
			{
				IStorage* pChildStg = nullptr;
				hr = pStorage->OpenStorage(pStatStg->pwcsName, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pChildStg);
				if (!SUCCEEDED(hr) || !pChildStg)
				{
					ti.m_hResult = hr;
					return;
				}
				// Recurse and traverse the child storage's elements
				LocStg::Traverse(pChildStg, ti);
				
				// If there was a problem below here, error out now
				if (!SUCCEEDED(ti.m_hResult)) return;
			}
			else if (STGTY_STREAM == pStatStg->type)
			{
				++(ti.m_nStreams);
				ti.m_cbTotalStreamSize.LowPart += pStatStg->cbSize.LowPart; // TODO: Allow for high part
			}
			else if (STGTY_LOCKBYTES == pStatStg->type)
			{
				++(ti.m_nLockbytes);
			}
			else if (STGTY_PROPERTY == pStatStg->type)
			{
				++(ti.m_nProperties);
			}
			++pStatStg;
		}

		// If there are no more blocks of children, return
		if (actualCount < BlockSize)
		{
			break;
		}
	}
}


HRESULT 
LocStg::OpenRootStorage(const std::wstring& strName, IStorage** o_ppStorage, bool bSilent)
{
	HRESULT hr = E_INVALIDARG;
	DWORD grfmode = STGM_DIRECT | STGM_READ | STGM_SHARE_DENY_WRITE;
	STGOPTIONS stgOptions;
	stgOptions.usVersion = 1;  // According to MSDN, 2 should be possible, but it isn't allowed (returns FAILED)
	stgOptions.reserved = 0;
	stgOptions.ulSectorSize = 0;

	hr = ::StgOpenStorageEx(strName.c_str(), grfmode, STGFMT_DOCFILE, 0, &stgOptions, 0, IID_IStorage, (void**)o_ppStorage);

	if (!bSilent)
	{
		if (SUCCEEDED(hr))
		{
			std::wcout << L"Storage options version = " << stgOptions.usVersion << std::endl;
			std::wcout << L"Sector size = " << stgOptions.ulSectorSize << std::endl;
		}
		else
		{
			std::wcout << L"Operation failed with status code " << hr << std::endl;
		}
	}
	return hr;
}

HRESULT
LocStg::CreateRootStorage(const std::wstring& strName, bool bOverwrite, IStorage** o_ppStorage)
{
	DWORD grfmode = STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE;
	grfmode |= bOverwrite ? STGM_CREATE : STGM_FAILIFTHERE;
	STGOPTIONS stgOptions;
	stgOptions.usVersion = 1;
	stgOptions.reserved = 0;
	stgOptions.ulSectorSize = 4096;

	return ::StgCreateStorageEx(strName.c_str(), grfmode, STGFMT_DOCFILE, 0, &stgOptions, 0, IID_IStorage, (void**)o_ppStorage);
}


HRESULT LocStg::GetStorageElementCount(IStorage* pStorage, int& elemCnt)
{
	IEnumSTATSTG *pEnum = nullptr;
	elemCnt = 0;
	HRESULT hr = pStorage->EnumElements(0, nullptr, 0, &pEnum);
	if (!SUCCEEDED(hr) || !pEnum)
	{
		std::wcout << L"Could not get a storage enumerator, error code = " << hr << std::endl;
		return hr;
	}
	ObjWrapper<IEnumSTATSTG> w;
	w.Set(pEnum);

	constexpr ULONG BlockSize = 100;
	STATSTG arrayStg[BlockSize];
	ULONG actualCount = 0;
	while (true)
	{
		hr = pEnum->Next(BlockSize, arrayStg, &actualCount);
		if (!(SUCCEEDED(hr)))
		{
			break;
		}
		elemCnt += actualCount;
		if (actualCount < BlockSize)
		{
			break;
		}
	}
	return hr;
}

HRESULT LocStg::FindLocalDocsStorage(IStorage *pParent, IStorage **o_ppLocalDocsStg)
{
	IEnumSTATSTG *pEnum = nullptr;
	*o_ppLocalDocsStg = nullptr;

	HRESULT hr = pParent->EnumElements(0, nullptr, 0, &pEnum);
	if (!SUCCEEDED(hr) || !pEnum)
	{
		std::wcout << L"Could not get a storage enumerator, error code = " << hr << std::endl;
		return hr;
	}
	ObjWrapper<IEnumSTATSTG> w;
	w.Set(pEnum);

	// Assumption: LocalDocs is in the first BlockSize of whatever storage we're searching!
	// It should be the 2nd storage within the root parent.
	constexpr ULONG BlockSize = 100;
	STATSTG arrayStg[BlockSize];
	ULONG actualCount = 0;

	hr = pEnum->Next(BlockSize, arrayStg, &actualCount);
	if (!(SUCCEEDED(hr)))
	{
		return hr;
	}
	STATSTG *pStatStg = &arrayStg[0];
	for (ULONG ii = 0; ii < actualCount; ++ii)
	{
		if (!(wcscmp(LocStg::sLocalDocs, pStatStg->pwcsName)))
		{
			// Found it in the info array, now open it and return the storage pointer
			hr = pParent->OpenStorage(pStatStg->pwcsName, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, o_ppLocalDocsStg);
			break;
		}
		++pStatStg;
	}
	return hr;
}

HRESULT LocStg::ListImmediateChildren(IStorage *pParent, int nChildren, int indentLevel)
{
	// If the caller supplied the number of children, use that. Otherwise get the count.

	int elementCount = nChildren;
	if (elementCount == 0)
	{
		HRESULT hr = LocStg::GetStorageElementCount(pParent, elementCount);
		if (!SUCCEEDED(hr))
		{
			std::wcout << L"GetStorageElementCount failed with error code " << hr << std::endl;
			return hr;
		}
	}

	IEnumSTATSTG *pEnum = nullptr;
	HRESULT hr = pParent->EnumElements(0, nullptr, 0, &pEnum);
	if (!SUCCEEDED(hr) || !pEnum)
	{
		std::wcout << L"Could not get a storage enumerator, error code = " << hr << std::endl;
		return hr;
	}

	std::unique_ptr<STATSTG[]> pArrayStg(new STATSTG[elementCount]);
	ULONG actualCount = 0;
	STATSTG *pStatStg = pArrayStg.get();
	hr = pEnum->Next(elementCount, pStatStg, &actualCount);
	if (!(SUCCEEDED(hr)))
	{
		return hr;
	}

	for (ULONG ii = 0; ii < actualCount; ++ii, ++pStatStg)
	{
		if (STGTY_STORAGE == pStatStg->type)
		{
			IStorage* pChildStg = nullptr;
			hr = pParent->OpenStorage(pStatStg->pwcsName, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pChildStg);
			if (SUCCEEDED(hr) && pChildStg)
			{
				_WriteIndent(indentLevel + 1);
				{
					int childCount = 0;
					HRESULT hr = LocStg::GetStorageElementCount(pChildStg, childCount);
					_WriteStorageType(pStatStg->type);
					std::wcout << L" " << pStatStg->pwcsName;
					if (SUCCEEDED(hr))
					{
						std::wcout << L" has " << childCount << L" elements" << std::endl;
					}
					else
					{
						std::wcout << L" has [unknown] elements" << std::endl;
					}
				}
				pChildStg->Release();
			}
		}
		else if (STGTY_STREAM == pStatStg->type)
		{
			LocStg::_WriteIndent(indentLevel + 1);
			LocStg::_WriteStorageType(pStatStg->type);
			std::wcout << L" " << pStatStg->pwcsName << L" size: (" << std::dec << (int)(pStatStg->cbSize.HighPart) << "),  " << std::dec << (int)(pStatStg->cbSize.LowPart) << std::endl;
		}
	}

	pEnum->Release();
	return hr;
}

void LocStg::_Usage(const wchar_t* thisProg)
{
	std::wcout << L"Usage: " << thisProg << L" <filename> [opcode]" << std::endl;
	std::wcout << L"   opcode:" << std::endl;
	std::wcout << L"      1: List summary information at top level" << std::endl;
	std::wcout << L"      2: Display total storage/stream counts" << std::endl;
}

void LocStg::_WaitForKey()
{
	std::cout << "Press return to continue" << std::endl;
	std::cin.get();
}

void LocStg::_WriteIndent(int nLevels)
{
	while (nLevels-- > 0)
	{
		std::wcout << L"    ";
	}
}

void LocStg::_WriteStorageType(DWORD type)
{
	switch (type) {
	case STGTY_STORAGE: std::wcout << L"Storage"; break;
	case STGTY_STREAM: std::wcout << L"Stream"; break;
	case STGTY_LOCKBYTES: std::wcout << L"Lockbytes"; break;
	case STGTY_PROPERTY: std::wcout << L"Property"; break;
	default: std::wcout << L"Unknown"; break;
	}
}
