#pragma once

#include <Windows.h>
#include <stdbool.h>

void InitializeCloaking();
void CleanupCloaking();
void tileWindows();
void toggleFullscreenMode();
void focusNextWindow(bool, unsigned int);
void gotoWorkspace(int);
void moveWindowToWorkspace(int);
void toggleDisableEnableTiling();
