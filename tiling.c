#include "tiling.h"
#include "error.h"
#include <Windows.h>
#include <objbase.h>
#include <shobjidl.h>
#include <winstring.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "windowsapp.lib")

// IApplicationView interface for window cloaking
typedef enum {
    AVCT_NONE = 0,
    AVCT_DEFAULT = 1,
    AVCT_VIRTUAL_DESKTOP = 2
} APPLICATION_VIEW_CLOAK_TYPE;

// IApplicationView interface GUID
static const GUID IID_IApplicationView = {0x372E1D3B, 0x38D3, 0x42E4, {0xA1, 0x5B, 0x8A, 0xB2, 0xB1, 0x78, 0xF5, 0x13}};

typedef struct IApplicationViewVtbl {
    HRESULT (*QueryInterface)(void* this, REFIID riid, void** ppvObject);
    ULONG (*AddRef)(void* this);
    ULONG (*Release)(void* this);
    HRESULT (*GetIids)(void* this, ULONG* iidCount, IID** iids);
    HRESULT (*GetRuntimeClassName)(void* this, void* className);
    HRESULT (*GetTrustLevel)(void* this, int* trustLevel);
    HRESULT (*SetFocusView)(void* this);
    HRESULT (*SwitchTo)(void* this);
    HRESULT (*TryInvokeBack)(void* this, void* callback);
    HRESULT (*GetThumbnailWindow)(void* this, HWND* hwnd);
    HRESULT (*GetMonitor)(void* this, void** monitor);
    HRESULT (*GetVisibility)(void* this, int* visibility);
    HRESULT (*SetCloak)(void* this, APPLICATION_VIEW_CLOAK_TYPE cloakType, int unknown);
} IApplicationViewVtbl;

struct IApplicationView {
    IApplicationViewVtbl* lpVtbl;
};

// IApplicationViewCollection interface GUID
static const GUID IID_IApplicationViewCollection = {0x1841C6D7, 0x4F9D, 0x42C0, {0xAF, 0x41, 0x87, 0x47, 0x53, 0x8F, 0x10, 0xE5}};

typedef struct IApplicationViewCollectionVtbl {
    HRESULT (*QueryInterface)(void* this, REFIID riid, void** ppvObject);
    ULONG (*AddRef)(void* this);
    ULONG (*Release)(void* this);
    HRESULT (*GetIids)(void* this, ULONG* iidCount, IID** iids);
    HRESULT (*GetRuntimeClassName)(void* this, void* className);
    HRESULT (*GetTrustLevel)(void* this, int* trustLevel);
    HRESULT (*GetViews)(void* this, void** array);
    HRESULT (*GetViewsByZOrder)(void* this, void** array);
    HRESULT (*GetViewsByAppUserModelId)(void* this, PCWSTR id, void** array);
    HRESULT (*GetViewForHwnd)(void* this, HWND hwnd, struct IApplicationView** view);
    HRESULT (*GetViewForApplication)(void* this, void* application, struct IApplicationView** view);
    HRESULT (*GetViewInFocus)(void* this, struct IApplicationView** view);
    HRESULT (*TryGetViewInFocus)(void* this, struct IApplicationView** view);
} IApplicationViewCollectionVtbl;

struct IApplicationViewCollection {
    IApplicationViewCollectionVtbl* lpVtbl;
};

// IServiceProvider interface GUID
static const GUID IID_IServiceProvider = {0x6D5140C1, 0x7436, 0x11CE, {0x80, 0x34, 0x00, 0xAA, 0x00, 0x60, 0x09, 0x34}};

// ImmersiveShell service CLSID
static const GUID CLSID_ImmersiveShell = {0xC2F03A33, 0x21F5, 0x47FA, {0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39}};

// Initialize COM for the application
void InitializeCloaking() {
    CoInitialize(NULL);
}

// Cleanup COM
void CleanupCloaking() {
    CoUninitialize();
}

// Function to cloak/uncloak window using IApplicationView
void SetWindowCloak(HWND hwnd, BOOL cloak) {
    IServiceProvider* serviceProvider = NULL;
    struct IApplicationViewCollection* viewCollection = NULL;
    struct IApplicationView* appView = NULL;
    
    HRESULT hr = CoCreateInstance(&CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER, &IID_IServiceProvider, (void**)&serviceProvider);
    if (FAILED(hr) || !serviceProvider) return;
    
    hr = serviceProvider->lpVtbl->QueryService(serviceProvider, &IID_IApplicationViewCollection, &IID_IApplicationViewCollection, (void**)&viewCollection);
    if (FAILED(hr) || !viewCollection) {
        serviceProvider->lpVtbl->Release(serviceProvider);
        return;
    }
    
    hr = viewCollection->lpVtbl->GetViewForHwnd(viewCollection, hwnd, &appView);
    if (SUCCEEDED(hr) && appView) {
        APPLICATION_VIEW_CLOAK_TYPE cloakType = cloak ? AVCT_VIRTUAL_DESKTOP : AVCT_NONE;
        appView->lpVtbl->SetCloak(appView, cloakType, 0);
        appView->lpVtbl->Release(appView);
    }
    
    viewCollection->lpVtbl->Release(viewCollection);
    serviceProvider->lpVtbl->Release(serviceProvider);
}

#define MAX_MANAGED 1024

typedef struct {
	HWND handle;
	int workspaceNumber;
	bool shouldCleanup;
} ManagedWindow;

bool isFullscreen = false;
bool isTilingEnabled = true;
HWND managed[MAX_MANAGED];
ManagedWindow totalManaged[MAX_MANAGED];
int numOfTotalManaged = 0;
int numOfCurrentlyManaged = 0;
int currentFocusedWindowIndex = 0;
int currentWorkspace = 1;
bool newWorkspace = false;

ManagedWindow* searchManaged(HWND handle)
{
	for (int i = 0; i < numOfTotalManaged; i++) {
		if (totalManaged[i].handle == handle) {
			return &totalManaged[i];
		}
	}

	return NULL;
}

void cleanupWorkspaceWindows()
{
	int keepCounter = 0;
	for (int i = 0; i < numOfTotalManaged; i++) {
		if (totalManaged[i].workspaceNumber == currentWorkspace) {
			continue;
		}

		totalManaged[keepCounter] = totalManaged[i];
		keepCounter++;
	}

	numOfTotalManaged = keepCounter;
}

BOOL isWindowManagable(HWND windowHandle)
{
	if (!IsWindowVisible(windowHandle) || IsHungAppWindow(windowHandle)) {
		return FALSE;
	}

	WINDOWINFO winInfo;
	winInfo.cbSize = sizeof(WINDOWINFO);
	if (!GetWindowInfo(windowHandle, &winInfo)) {
		return FALSE;
	}

	if (winInfo.dwStyle & WS_POPUP) {
		return FALSE;
	}

	if (!(winInfo.dwExStyle & 0x20000000)) {
		return FALSE;
	}

	if (GetWindowTextLengthW(windowHandle) == 0) {
		return FALSE;
	}

	RECT clientRect;
	if (!GetClientRect(windowHandle, &clientRect)) {
		return FALSE;
	}

	// Skip small windows to avoid bugs
	if (clientRect.right < 100 || clientRect.bottom < 100){
		return FALSE;
	}

	return TRUE;
}

BOOL CALLBACK windowEnumeratorCallback(HWND currentWindowHandle, LPARAM lparam)
{
	if (numOfTotalManaged > MAX_MANAGED) {
		return FALSE;
	}

	if (!isWindowManagable(currentWindowHandle)) {
		return TRUE;
	}

	if (searchManaged(currentWindowHandle) != NULL) {
		return TRUE;
	}

	totalManaged[numOfTotalManaged].handle = currentWindowHandle;
	totalManaged[numOfTotalManaged].workspaceNumber = currentWorkspace;
	totalManaged[numOfTotalManaged].shouldCleanup = false;
	numOfTotalManaged++;

	return TRUE;
}

void updateManagedWindows()
{
	numOfCurrentlyManaged = 0;

	if (isFullscreen) {
		managed[0] = GetForegroundWindow();
		numOfCurrentlyManaged = 1;
		return;
	}

	// First pass: hide all windows not in current workspace
	for (int i = 0; i < numOfTotalManaged; i++) {
		if (totalManaged[i].workspaceNumber != currentWorkspace) {
			// Use IApplicationView SetCloak to hide window instantly without animation
			SetWindowCloak(totalManaged[i].handle, TRUE);
		}
	}

	// Second pass: show and manage windows in current workspace
	for (int i = 0; i < numOfTotalManaged; i++) {
		if (totalManaged[i].workspaceNumber != currentWorkspace) {
			continue;
		}

		managed[numOfCurrentlyManaged] = totalManaged[i].handle;
		// Use IApplicationView SetCloak to show window instantly without animation
		SetWindowCloak(managed[numOfCurrentlyManaged], FALSE);
		numOfCurrentlyManaged++;
	}
}

void tileWindows()
{
	if (!isTilingEnabled) {
		return;
	}

	if (newWorkspace) {
		newWorkspace = false;
	} else {
		if (!isFullscreen) {
			cleanupWorkspaceWindows();
		}
	}

	EnumChildWindows(GetDesktopWindow(), windowEnumeratorCallback, 0);

	if (numOfTotalManaged == 0) {
		return;
	}

	updateManagedWindows();

	TileWindows(GetDesktopWindow(), MDITILE_VERTICAL | MDITILE_SKIPDISABLED, NULL, numOfCurrentlyManaged, managed);
}

void toggleFullscreenMode()
{
	isFullscreen = !isFullscreen;
	newWorkspace = true;
	tileWindows();
}

void focusNextWindow(bool goBack, unsigned int callCount)
{
	// Avoid infinite recursion
	if (callCount > 25) {
		tileWindows();
		return;
	}

	if (isFullscreen) {
		toggleFullscreenMode();
	}

	currentFocusedWindowIndex += goBack ? -1 : 1;

	if (currentFocusedWindowIndex < 0) {
		currentFocusedWindowIndex = numOfCurrentlyManaged - 1;
	} else if (currentFocusedWindowIndex >= numOfCurrentlyManaged) {
		currentFocusedWindowIndex = 0;
	}

	if (!isWindowManagable(managed[currentFocusedWindowIndex]) || (GetForegroundWindow() == managed[currentFocusedWindowIndex])) {
		focusNextWindow(goBack, ++callCount);
	}

	SwitchToThisWindow(managed[currentFocusedWindowIndex], FALSE);
}

void gotoWorkspace(int number)
{
	if (isFullscreen) {
		isFullscreen = !isFullscreen;
		newWorkspace = true;
	}

	tileWindows();

	currentWorkspace = number;
	newWorkspace = true;
	tileWindows();
}

void moveWindowToWorkspace(int workspaceNumber)
{
	if (numOfCurrentlyManaged == 0 || workspaceNumber == currentWorkspace) {
		return;
	}

	ManagedWindow* managedWindow = searchManaged(GetForegroundWindow());
	if (managedWindow == NULL) {
		return;
	}

	// Use IApplicationView SetCloak to hide window instantly without animation
	SetWindowCloak(managedWindow->handle, TRUE);
	managedWindow->workspaceNumber = workspaceNumber;
	tileWindows();
}

void toggleDisableEnableTiling() {
	isTilingEnabled = !isTilingEnabled;

	if (isTilingEnabled) {
		tileWindows();
	}
}
