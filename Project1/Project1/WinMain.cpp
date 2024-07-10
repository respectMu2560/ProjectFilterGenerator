#include <windows.h>
#include <lmcons.h>
#include <commctrl.h>
#include <windowsx.h>
#include <shlobj.h>
#include <objbase.h>
#include <shobjidl.h>
#include <filesystem>
#include <string>
#include <vector>
#include <map>
#include <fstream>
#include <set>
#include <algorithm>
#include <regex>
#include <objbase.h>
#include <codecvt>
#include <locale>
#include <system_error>
#include <iomanip>
#include <sstream>
#include <numeric>
#include "tinyxml2.h"
#include "resource.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(linker, "/SUBSYSTEM:WINDOWS /ENTRY:wWinMainCRTStartup")

#define IDC_TREEVIEW 1001
#define IDC_OK_BUTTON 1002
#define IDC_CANCEL_BUTTON 1003

namespace fs = std::filesystem;

std::string wstring_to_utf8(const std::wstring& wstr);

// 前回のコードから再利用する構造体と関数
struct FileInfo {
	std::wstring name;
	std::wstring filter;
};

std::wstring GetCurrentUserName() {
	wchar_t username[UNLEN + 1];
	DWORD username_len = UNLEN + 1;
	if (GetUserNameW(username, &username_len)) {
		return std::wstring(username);
	}
	return L"Unknown";
}

std::wstring generateGuid() {
	GUID guid;
	HRESULT hr = CoCreateGuid(&guid);

	if (FAILED(hr)) {
		throw std::runtime_error("Failed to create GUID");
	}

	wchar_t guidString[39];
	int result = StringFromGUID2(guid, guidString, sizeof(guidString) / sizeof(wchar_t));

	if (result == 0) {
		throw std::runtime_error("Failed to convert GUID to string");
	}

	// StringFromGUID2 returns the GUID in the format {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
	// We need to remove the curly braces for the .vcxproj.filters file format
	std::wstring guidStr(guidString);
	return guidStr.substr(1, guidStr.length() - 2);
}

std::vector<FileInfo> getProjectFiles(const fs::path& directory) {
	std::vector<FileInfo> files;
	std::wstring baseDir = directory.wstring();

	for (const auto& entry : fs::recursive_directory_iterator(directory)) {
		if (entry.is_regular_file()) {
			std::wstring ext = entry.path().extension().wstring();
			std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

			// C++ソースファイルとヘッダーファイルのみを対象とする
			if (ext == L".cpp" || ext == L".h" || ext == L".hpp" || ext == L".c") {
				std::wstring relativePath = fs::relative(entry.path(), directory).wstring();
				std::wstring filter = fs::relative(entry.path().parent_path(), directory).wstring();

				// ルートディレクトリの場合、filterを空文字列にする
				if (filter == L".") {
					filter = L"";
				}

				// Windowsのパス区切り文字をスラッシュに変換
				std::replace(relativePath.begin(), relativePath.end(), L'\\', L'/');
				std::replace(filter.begin(), filter.end(), L'\\', L'/');

				files.push_back({ relativePath, filter });
			}
		}
	}

	return files;
}

std::wstring xmlEscape(const std::wstring& input) {
	std::wstring escaped;
	for (wchar_t ch : input) {
		switch (ch) {
		case L'&':  escaped += L"&amp;"; break;
		case L'<':  escaped += L"&lt;"; break;
		case L'>':  escaped += L"&gt;"; break;
		case L'"':  escaped += L"&quot;"; break;
		case L'\'': escaped += L"&apos;"; break;
		default:    escaped += ch; break;
		}
	}
	return escaped;
}

void generateFiltersFile(const std::wstring& projectDirectory, const std::wstring& projectName, const std::vector<FileInfo>& files) {
	fs::path filtersPath = fs::path(projectDirectory) / (std::wstring(projectName) + L".vcxproj.filters");

	try {
		// ディレクトリの存在確認と作成、既存ファイルの削除、ディスク容量確認は以前と同じ

		std::wofstream filtersFile(filtersPath, std::ios::out | std::ios::trunc);
		if (!filtersFile.is_open()) {
			throw std::runtime_error("Failed to open filters file: " + wstring_to_utf8(filtersPath.wstring()));
		}

		filtersFile << L"<?xml version=\"1.0\" encoding=\"utf-8\"?>\n";
		filtersFile << L"<Project ToolsVersion=\"4.0\" xmlns=\"http://schemas.microsoft.com/developer/msbuild/2003\">\n";

		// フィルター（ディレクトリ）の階層構造を構築
		std::map<std::wstring, std::wstring> filterHierarchy;
		for (const auto& file : files) {
			std::wstring filePath = file.name;
			std::replace(filePath.begin(), filePath.end(), L'/', L'\\');  // '/' を '\' に置換

			fs::path fsFilePath(filePath);
			std::wstring currentPath;
			for (const auto& part : fsFilePath.parent_path()) {
				if (!currentPath.empty()) {
					std::wstring parentPath = currentPath;
					currentPath += L"\\" + part.wstring();
					if (filterHierarchy.find(currentPath) == filterHierarchy.end()) {
						filterHierarchy[currentPath] = parentPath;
					}
				}
				else {
					currentPath = part.wstring();
					if (filterHierarchy.find(currentPath) == filterHierarchy.end()) {
						filterHierarchy[currentPath] = L"";
					}
				}
			}
		}

		// フィルターの書き込み
		filtersFile << L"  <ItemGroup>\n";
		for (const auto& [filter, parentFilter] : filterHierarchy) {
			filtersFile << L"    <Filter Include=\"" << filter << L"\">\n";
			if (!parentFilter.empty()) {
				filtersFile << L"      <Filter>" << parentFilter << L"</Filter>\n";
			}
			filtersFile << L"      <UniqueIdentifier>{" << generateGuid() << L"}</UniqueIdentifier>\n";
			filtersFile << L"    </Filter>\n";
		}
		filtersFile << L"  </ItemGroup>\n";

		// ファイルの書き込み
		filtersFile << L"  <ItemGroup>\n";
		for (const auto& file : files) {
			std::wstring filePath = file.name;
			std::replace(filePath.begin(), filePath.end(), L'/', L'\\');  // '/' を '\' に置換

			fs::path fsFilePath(filePath);
			std::wstring ext = fsFilePath.extension().wstring();
			std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

			std::wstring itemType;
			if (ext == L".cpp" || ext == L".c") {
				itemType = L"ClCompile";
			}
			else if (ext == L".h" || ext == L".hpp") {
				itemType = L"ClInclude";
			}
			else {
				itemType = L"None";
			}

			filtersFile << L"    <" << itemType << L" Include=\"" << filePath << L"\">\n";

			std::wstring filterPath = fsFilePath.parent_path().wstring();
			if (!filterPath.empty()) {
				filtersFile << L"      <Filter>" << filterPath << L"</Filter>\n";
			}

			filtersFile << L"    </" << itemType << L">\n";
		}
		filtersFile << L"  </ItemGroup>\n";

		filtersFile << L"</Project>\n";

		filtersFile.close();
		if (filtersFile.fail()) {
			throw std::runtime_error("Failed to close filters file");
		}

		// ファイルが正常に作成されたか確認
		if (!fs::exists(filtersPath) || fs::file_size(filtersPath) == 0) {
			throw std::runtime_error("Filters file was not created successfully or is empty");
		}
	}
	catch (const std::exception& e) {
		std::stringstream ss;
		ss << "Error in generateFiltersFile: " << e.what() << "\n";
		ss << "Partial file contents:\n";
		std::wifstream partialFile(filtersPath);
		if (partialFile.is_open()) {
			std::wstring line;
			for (int i = 0; i < 20 && std::getline(partialFile, line); ++i) {
				ss << wstring_to_utf8(line) << "\n";
			}
			partialFile.close();
		}
		throw std::runtime_error(ss.str());
	}
}

HWND hWndDirectoryEdit, hWndProjectNameEdit, hWndGenerateButton, hWndStatusText, hWndBrowseButton;
WNDPROC oldEditProc;
HFONT hFont;
HWND hTreeView;
HWND hConfirmDialog;

// フォント作成関数
HFONT CreateCustomFont(int size = 16) {
	return CreateFont(
		size, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
		DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
		CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
		L"Segoe UI"  // ここでフォント名を指定
	);
}

// プロジェクトファイルを検索し、プロジェクト名を抽出する関数
std::wstring FindProjectFile(const std::wstring& directory) {
	for (const auto& entry : fs::directory_iterator(directory)) {
		if (entry.path().extension() == L".vcxproj") {
			std::wifstream file(entry.path());
			std::wstring line;
			std::wregex projectNameRegex(L"<ProjectName>(.+?)</ProjectName>");
			while (std::getline(file, line)) {
				std::wsmatch match;
				if (std::regex_search(line, match, projectNameRegex)) {
					return match[1].str();
				}
			}
			// ProjectName要素が見つからない場合、ファイル名（拡張子なし）を返す
			return entry.path().stem().wstring();
		}
	}
	return L"";  // プロジェクトファイルが見つからない場合
}

// wstringをstringに変換するユーティリティ関数
std::string ws2s(const std::wstring& wstr) {
	std::string str(wstr.begin(), wstr.end());
	return str;
}

std::string wstring_to_utf8(const std::wstring& wstr) {
	if (wstr.empty()) return std::string();
	int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
	std::string strTo(size_needed, 0);
	WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
	return strTo;
}

void updateProjectFile(const std::wstring& projectDirectory, const std::wstring& projectName, const std::vector<FileInfo>& files) {
	fs::path projectPath = fs::path(projectDirectory) / (std::wstring(projectName) + L".vcxproj");
	fs::path backupPath = projectPath.parent_path() / (projectPath.stem().wstring() + L"_backup.vcxproj");

	try {
		// プロジェクトファイルのバックアップを作成
		fs::copy_file(projectPath, backupPath, fs::copy_options::overwrite_existing);

		// プロジェクトファイルを読み込む
		std::wifstream inFile(projectPath);
		if (!inFile) {
			throw std::runtime_error("Failed to open project file for reading");
		}

		std::vector<std::wstring> lines;
		std::wstring line;
		while (std::getline(inFile, line)) {
			lines.push_back(line);
		}
		inFile.close();

		// 既存のファイルエントリを確認
		std::map<std::wstring, std::pair<std::wstring, size_t>> existingFiles; // <ファイル名, <アイテムタイプ, 行番号>>
		std::wregex fileRegex(L"<(ClCompile|ClInclude)\\s+Include=\"([^\"]+)\"");

		for (size_t i = 0; i < lines.size(); ++i) {
			std::wsmatch fileMatch;
			if (std::regex_search(lines[i], fileMatch, fileRegex)) {
				std::wstring itemType = fileMatch[1].str();
				std::wstring fileName = fileMatch[2].str();
				existingFiles[fileName] = std::make_pair(itemType, i);
			}
		}

		// 新しいファイルエントリを準備
		std::vector<std::wstring> newEntries;
		for (const auto& file : files) {
			std::wstring normalizedPath = file.name;
			std::replace(normalizedPath.begin(), normalizedPath.end(), L'/', L'\\');

			std::wstring ext = fs::path(normalizedPath).extension().wstring();
			std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

			std::wstring itemType = (ext == L".cpp" || ext == L".c") ? L"ClCompile" : L"ClInclude";

			auto it = existingFiles.find(normalizedPath);
			if (it == existingFiles.end() || it->second.first != itemType) {
				std::wstring newEntry = L"    <" + itemType + L" Include=\"" + normalizedPath + L"\" />";
				newEntries.push_back(newEntry);
			}
		}

		// 新しいエントリを適切な位置に挿入
		if (!newEntries.empty()) {
			auto insertPos = std::find_if(lines.rbegin(), lines.rend(), [](const std::wstring& l) {
				return l.find(L"</ItemGroup>") != std::wstring::npos;
				});

			if (insertPos != lines.rend()) {
				size_t insertIndex = lines.size() - (insertPos - lines.rbegin()) - 1;
				lines.insert(lines.begin() + insertIndex, newEntries.begin(), newEntries.end());
			}
			else {
				// ItemGroupが見つからない場合、ファイルの最後に新しいItemGroupを追加
				lines.push_back(L"  <ItemGroup>");
				lines.insert(lines.end(), newEntries.begin(), newEntries.end());
				lines.push_back(L"  </ItemGroup>");
			}
		}

		// 更新されたコンテンツを書き込む
		std::wofstream outFile(projectPath);
		if (!outFile) {
			throw std::runtime_error("Failed to open project file for writing");
		}
		for (const auto& line : lines) {
			outFile << line << L"\n";
		}
		outFile.close();

		// バックアップを削除
		fs::remove(backupPath);
	}
	catch (const std::exception& e) {
		// エラーが発生した場合、バックアップから復元
		if (fs::exists(backupPath)) {
			fs::copy_file(backupPath, projectPath, fs::copy_options::overwrite_existing);
			fs::remove(backupPath);
		}
		throw std::runtime_error(std::string("Error in updateProjectFile: ") + e.what());
	}
}

// ディレクトリパスが変更されたときにプロジェクト名を自動更新する関数
void UpdateProjectName(HWND hwnd) {
	wchar_t directoryPath[MAX_PATH];
	GetWindowText(hWndDirectoryEdit, directoryPath, MAX_PATH);

	// ディレクトリの存在確認
	if (!fs::is_directory(directoryPath)) {
		SetWindowText(hWndProjectNameEdit, L"");
		SetWindowText(hWndStatusText, L"Invalid directory path.");
		return;
	}

	std::wstring projectName = FindProjectFile(directoryPath);
	if (!projectName.empty()) {
		SetWindowText(hWndProjectNameEdit, projectName.c_str());
		SetWindowText(hWndStatusText, L"");  // ステータスメッセージをクリア
	}
	else {
		SetWindowText(hWndProjectNameEdit, L"");
		SetWindowText(hWndStatusText, L"No .vcxproj file found in the selected directory.");
	}
}


// ファイル選択ダイアログを表示する関数
std::wstring BrowseFolder() {
	IFileOpenDialog* pFileOpen;
	std::wstring selectedPath;

	// Create the FileOpenDialog object
	HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL,
		IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));

	if (SUCCEEDED(hr)) {
		// Set options
		DWORD dwOptions;
		pFileOpen->GetOptions(&dwOptions);
		pFileOpen->SetOptions(dwOptions | FOS_PICKFOLDERS);

		// Show the Open dialog box
		hr = pFileOpen->Show(NULL);

		if (SUCCEEDED(hr)) {
			IShellItem* pItem;
			hr = pFileOpen->GetResult(&pItem);
			if (SUCCEEDED(hr)) {
				PWSTR pszFilePath;
				hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
				if (SUCCEEDED(hr)) {
					selectedPath = pszFilePath;
					CoTaskMemFree(pszFilePath);
				}
				pItem->Release();
			}
		}
		pFileOpen->Release();
	}
	return selectedPath;
}

struct FilterNode {
	std::wstring name;
	std::vector<FilterNode> children;
	bool isFile;  // ファイルかどうかを示すフラグ
};

FilterNode buildFilterStructure(const std::wstring& directory) {
	FilterNode root;
	root.name = fs::path(directory).filename().wstring();
	root.isFile = false;

	for (const auto& entry : fs::recursive_directory_iterator(directory)) {
		std::wstring ext = entry.path().extension().wstring();
		std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

		if (entry.is_regular_file() && (ext == L".cpp" || ext == L".h" || ext == L".hpp")) {
			std::vector<std::wstring> path_parts;
			for (const auto& part : entry.path().lexically_relative(directory)) {
				path_parts.push_back(part.wstring());
			}

			FilterNode* current = &root;
			for (size_t i = 0; i < path_parts.size(); ++i) {
				auto& part = path_parts[i];
				auto it = std::find_if(current->children.begin(), current->children.end(),
					[&part](const FilterNode& node) { return node.name == part; });

				if (it == current->children.end()) {
					current->children.emplace_back(FilterNode{ part, {}, i == path_parts.size() - 1 });
					current = &current->children.back();
				}
				else {
					current = &(*it);
				}
			}
		}
	}
	return root;
}

HTREEITEM AddItemToTree(HWND hTreeView, HTREEITEM hParent, const std::wstring& text) {
	TVINSERTSTRUCT tvInsert = { 0 };
	tvInsert.hParent = hParent;
	tvInsert.hInsertAfter = TVI_LAST;
	tvInsert.item.mask = TVIF_TEXT;
	tvInsert.item.pszText = const_cast<LPWSTR>(text.c_str());
	return TreeView_InsertItem(hTreeView, &tvInsert);
}

void PopulateTreeView(HWND hTreeView, HTREEITEM hParent, const FilterNode& node) {
	HTREEITEM hItem = AddItemToTree(hTreeView, hParent, node.name.c_str());
	for (const auto& child : node.children) {
		if (child.isFile) {
			// ファイルの場合、名前をそのまま表示
			AddItemToTree(hTreeView, hItem, child.name.c_str());
		}
		else {
			// ディレクトリの場合、再帰的に処理
			PopulateTreeView(hTreeView, hItem, child);
		}
	}
}

INT_PTR CALLBACK ConfirmDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
	switch (uMsg) {
	case WM_INITDIALOG:
	{
		FilterNode* root = reinterpret_cast<FilterNode*>(lParam);

		hTreeView = CreateWindowEx(0, WC_TREEVIEW, L"",
			WS_VISIBLE | WS_CHILD | WS_BORDER | TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS,
			10, 10, 380, 300, hwnd, (HMENU)IDC_TREEVIEW, GetModuleHandle(NULL), NULL);
		SendMessage(hTreeView, WM_SETFONT, (WPARAM)hFont, TRUE);

		PopulateTreeView(hTreeView, TVI_ROOT, *root);

		HWND hOkButton = CreateWindow(L"BUTTON", L"OK", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
			200, 320, 80, 30, hwnd, (HMENU)IDOK, GetModuleHandle(NULL), NULL);
		SendMessage(hOkButton, WM_SETFONT, (WPARAM)hFont, TRUE);

		HWND hCancelButton = CreateWindow(L"BUTTON", L"Cancel", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
			290, 320, 80, 30, hwnd, (HMENU)IDCANCEL, GetModuleHandle(NULL), NULL);
		SendMessage(hCancelButton, WM_SETFONT, (WPARAM)hFont, TRUE);
	}
	return TRUE;

	case WM_COMMAND:
		switch (LOWORD(wParam)) {
		case IDOK:
		case IDCANCEL:
			EndDialog(hwnd, LOWORD(wParam));
			return TRUE;
		}
		break;

	case WM_CLOSE:
		EndDialog(hwnd, IDCANCEL);
		return TRUE;
	}

	return FALSE;
}

// EditコントロールのサブクラスプロシージャWNDPROC
LRESULT CALLBACK EditSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
	switch (uMsg) {
	case WM_DROPFILES: {
		HDROP hDrop = (HDROP)wParam;
		wchar_t szFileName[MAX_PATH];
		if (DragQueryFile(hDrop, 0, szFileName, MAX_PATH)) {
			// ドロップされたのがディレクトリかチェック
			DWORD fileAttributes = GetFileAttributes(szFileName);
			if (fileAttributes != INVALID_FILE_ATTRIBUTES && (fileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
				SetWindowText(hwnd, szFileName);
				UpdateProjectName(GetParent(hwnd));
			}
			else {
				SetWindowText(hWndStatusText, L"Please drop a directory, not a file.");
			}
		}
		DragFinish(hDrop);
		return 0;
	}
	case WM_KEYUP:
	case WM_KILLFOCUS:
		UpdateProjectName(GetParent(hwnd));
		break;
	}
	return CallWindowProc(oldEditProc, hwnd, uMsg, wParam, lParam);
}

// ウィンドウプロシージャ
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
	switch (uMsg) {
	case WM_CREATE: {
		// カスタムフォントを作成
		hFont = CreateCustomFont();

		// ディレクトリパス入力欄
		HWND hWndStatic = CreateWindow(L"STATIC", L"Project Directory:", WS_VISIBLE | WS_CHILD,
			10, 10, 150, 20, hwnd, NULL, NULL, NULL);
		SendMessage(hWndStatic, WM_SETFONT, (WPARAM)hFont, TRUE);

		hWndDirectoryEdit = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
			10, 30, 250, 20, hwnd, NULL, NULL, NULL);
		SendMessage(hWndDirectoryEdit, WM_SETFONT, (WPARAM)hFont, TRUE);

		// ブラウズボタン
		hWndBrowseButton = CreateWindow(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
			270, 30, 60, 20, hwnd, (HMENU)2, NULL, NULL);
		SendMessage(hWndBrowseButton, WM_SETFONT, (WPARAM)hFont, TRUE);

		// プロジェクト名入力欄 (読み取り専用)
		HWND hWndStatic2 = CreateWindow(L"STATIC", L"Project Name:", WS_VISIBLE | WS_CHILD,
			10, 60, 150, 20, hwnd, NULL, NULL, NULL);
		SendMessage(hWndStatic2, WM_SETFONT, (WPARAM)hFont, TRUE);

		hWndProjectNameEdit = CreateWindow(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL | ES_READONLY,
			10, 80, 300, 20, hwnd, NULL, NULL, NULL);
		SendMessage(hWndProjectNameEdit, WM_SETFONT, (WPARAM)hFont, TRUE);

		// 生成ボタン
		hWndGenerateButton = CreateWindow(L"BUTTON", L"Generate Filters", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
			10, 110, 150, 30, hwnd, (HMENU)1, NULL, NULL);
		SendMessage(hWndGenerateButton, WM_SETFONT, (WPARAM)hFont, TRUE);

		// ステータステキスト
		hWndStatusText = CreateWindow(L"STATIC", L"", WS_VISIBLE | WS_CHILD,
			10, 150, 300, 20, hwnd, NULL, NULL, NULL);
		SendMessage(hWndStatusText, WM_SETFONT, (WPARAM)hFont, TRUE);

		// ドラッグアンドドロップを有効にする
		DragAcceptFiles(hWndDirectoryEdit, TRUE);

		// EditコントロールをサブクラSQ
		oldEditProc = (WNDPROC)SetWindowLongPtr(hWndDirectoryEdit, GWLP_WNDPROC, (LONG_PTR)EditSubclassProc);

		break;
	}
	case WM_COMMAND:
		if (LOWORD(wParam) == 1) {  // Generate Filters ボタンが押された
			wchar_t directoryPath[MAX_PATH], projectName[MAX_PATH];
			GetWindowText(hWndDirectoryEdit, directoryPath, MAX_PATH);
			GetWindowText(hWndProjectNameEdit, projectName, MAX_PATH);

			if (!fs::is_directory(directoryPath)) {
				SetWindowText(hWndStatusText, L"Invalid directory path. Please select a valid directory.");
				return 0;
			}

			if (wcslen(projectName) == 0) {
				SetWindowText(hWndStatusText, L"Project name is empty. Please select a valid directory with a .vcxproj file.");
				return 0;
			}

			FilterNode root = buildFilterStructure(directoryPath);

			// Show confirmation dialog
			INT_PTR result = DialogBoxParam(
				GetModuleHandle(NULL),
				MAKEINTRESOURCE(IDD_CONFIRM_DIALOG),
				hwnd,
				ConfirmDialogProc,
				reinterpret_cast<LPARAM>(&root)
			);

			if (result == IDOK) {
				auto files = getProjectFiles(directoryPath);
				try {
					generateFiltersFile(directoryPath, projectName, files);
					updateProjectFile(directoryPath, projectName, files);
					SetWindowText(hWndStatusText, L"Filters generated and project file updated successfully.");
				}
				catch (const std::exception& e) {
					std::stringstream ss;
					ss << "Error: " << e.what() << "\n";

					fs::path filtersPath = fs::path(directoryPath) / (std::wstring(projectName) + L".vcxproj.filters");
					if (fs::exists(filtersPath)) {
						try {
							fs::file_status s = fs::status(filtersPath);
							ss << "File permissions: " << static_cast<int>(s.permissions()) << "\n";

							// ファイルの内容を読み込んで確認
							std::wifstream readFile(filtersPath);
							if (readFile.is_open()) {
								ss << "File contents:\n";
								std::wstring line;
								while (std::getline(readFile, line)) {
									ss << wstring_to_utf8(line) << "\n";
								}
								readFile.close();
							}
							else {
								ss << "Unable to read file contents.\n";
							}
						}
						catch (const fs::filesystem_error& fe) {
							ss << "Error checking file status: " << fe.what() << "\n";
						}
					}
					else {
						ss << "Filters file does not exist.\n";
					}

					ss << "Project Directory: " << wstring_to_utf8(directoryPath) << "\n"
						<< "Project Name: " << wstring_to_utf8(projectName) << "\n"
						<< "Current Directory: " << wstring_to_utf8(fs::current_path().wstring()) << "\n"
						<< "Free Disk Space: " << fs::space(directoryPath).available / (1024 * 1024) << " MB\n"
						<< "User Name: " << wstring_to_utf8(GetCurrentUserName()) << "\n"
						<< "Process ID: " << GetCurrentProcessId();

					std::string errorMsg = ss.str();
					OutputDebugStringA(errorMsg.c_str());  // デバッグコンソールにエラーを出力

					// エラーメッセージをログファイルに書き込む（上書きモード）
					std::ofstream logFile("error_log.txt", std::ios::out | std::ios::trunc);
					if (logFile.is_open()) {
						logFile << "Error occurred at " << std::time(nullptr) << ":\n" << errorMsg << "\n";
						logFile.close();
					}

					std::wstring wideErrorMsg = std::wstring(errorMsg.begin(), errorMsg.end());
					SetWindowText(hWndStatusText, wideErrorMsg.c_str());

					// エラーメッセージボックスを表示
					MessageBoxA(hwnd, errorMsg.c_str(), "Error", MB_OK | MB_ICONERROR);
				}
			}
			else {
				SetWindowText(hWndStatusText, L"Filter generation cancelled.");
			}
		}
		else if (LOWORD(wParam) == 2) {  // Browse ボタンが押された
			std::wstring selectedPath = BrowseFolder();
			if (!selectedPath.empty()) {
				SetWindowText(hWndDirectoryEdit, selectedPath.c_str());
				UpdateProjectName(hwnd);
			}
		}
		break;
	case WM_DESTROY:
		DeleteObject(hFont);  // フォントリソースを解放
		PostQuitMessage(0);
		return 0;
	}
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
	// COMライブラリの初期化
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	if (FAILED(hr)) {
		MessageBox(NULL, L"COM initialization failed", L"Error", MB_OK | MB_ICONERROR);
		return 1;
	}

	// ウィンドウクラスを登録
	const wchar_t CLASS_NAME[] = L"Filter Generator Window Class";

	WNDCLASS wc = {};
	wc.lpfnWndProc = WindowProc;
	wc.hInstance = hInstance;
	wc.lpszClassName = CLASS_NAME;

	RegisterClass(&wc);

	HWND hwnd = CreateWindowEx(
		0,                              // オプションのウィンドウスタイル
		CLASS_NAME,                     // ウィンドウクラス
		L"VS Filter Generator",         // ウィンドウテキスト
		WS_OVERLAPPEDWINDOW,            // ウィンドウスタイル
		CW_USEDEFAULT, CW_USEDEFAULT,   // 位置
		400, 250,                       // サイズ
		NULL,                           // 親ウィンドウ    
		NULL,                           // メニュー
		hInstance,                      // インスタンスハンドル
		NULL                            // 追加のアプリケーションデータ
	);

	wc.lpfnWndProc = ConfirmDialogProc;
	wc.hInstance = hInstance;
	wc.lpszClassName = L"ConfirmDialogClass";
	RegisterClass(&wc);

	if (hwnd == NULL) {
		return 0;
	}

	// DPIを考慮してフォントサイズを調整
	HDC hdc = GetDC(hwnd);
	int dpi = GetDeviceCaps(hdc, LOGPIXELSY);
	ReleaseDC(hwnd, hdc);
	hFont = CreateCustomFont(-MulDiv(11, dpi, 72));  // 11ポイントのフォントを作成

	ShowWindow(hwnd, nCmdShow);

	// メッセージループ
	MSG msg = {};
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	// COMライブラリの解放
	CoUninitialize();

	return 0;
}