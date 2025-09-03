/*
	* ProcessAudioRecorder 入口模块
	* 功能概述：
	*   本模块实现了一个命令行音频录制工具的核心控制逻辑，支持三种进程级音频捕获模式：
	*     1. 全局录制（系统混音）
	*     2. 进程包含模式（仅录制指定进程）
	*     3. 进程排除模式（排除指定进程）
	* 实现机制：
	*   - 命令行解析：使用自定义参数解析器处理 --pid/--mode/--path 参数组合
	*   - 进程监控：实时检测目标进程状态，自动终止失效进程的录音
	*   - 异步控制：通过控制台事件处理器实现 Ctrl+C 优雅停止
	*   - 进度显示：动态刷新录制时长（“时:分:秒”格式）
	*   - 错误处理：全面捕获 COM 异常并转换为可读错误信息
	* 核心流程：
	*   1. 解析命令行参数 → 2. 验证目标进程状态 → 3. 初始化音频捕获引擎
	*   4. 启动后台录制线程 → 5. 监控停止条件 → 6. 优雅停止并保存文件
*/

#include <audioclient.h>
#include <chrono>
#include <atomic>
#include <iomanip>
#include <iostream>
#include <map>
#include <string>
#include <thread>
#include <Windows.h>
#include <locale.h>
#include "LoopbackCapture.h"

static std::atomic<bool> g_bStopCapture(false);

static CLoopbackCapture* g_pCurrentCapture = nullptr;

struct CommandLineArgs {
	DWORD processId = 0;
	int captureMode = 1;
	std::wstring outputPath;
	bool isValid = false;
	std::wstring errorMessage;
};

CommandLineArgs ParseCommandLine(int argc, wchar_t* argv[]) {
	CommandLineArgs args;
	if (argc < 2) {
		args.errorMessage = L"Error: Too few arguments.";
		return args;
	}
	std::map<std::wstring, std::wstring> params;
	for (int i = 1; i < argc; i++) {
		std::wstring arg = argv[i];
		if (arg.substr(0, 2) == L"--") {
			std::wstring option = arg.substr(2);
			if (i + 1 < argc && argv[i + 1][0] != L'-') {
				params[option] = argv[i + 1];
				i++;
			}
			else {
				params[option] = L"";
			}
		}
		else {
			args.errorMessage = L"Error: Invalid argument format: " + arg +
				L"\nAll arguments must be options starting with --";
			return args;
		}
	}

	if (params.find(L"path") == params.end()) {
		args.errorMessage = L"Error: Missing required argument --path";
		return args;
	}
	args.outputPath = params[L"path"];
	if (args.outputPath.empty()) {
		args.errorMessage = L"Error: Output path cannot be empty.";
		return args;
	}

	wchar_t* endPtr;
	if (params.find(L"mode") != params.end()) {
		args.captureMode = std::wcstol(params[L"mode"].c_str(), &endPtr, 10);
		if (endPtr == params[L"mode"].c_str() || *endPtr != L'\0' ||
			(args.captureMode != 0 && args.captureMode != 1 && args.captureMode != 2)) {
			args.errorMessage = L"Error: Invalid mode: " + params[L"mode"] +
				L"\nMust be 0 (global), 1 (include), or 2 (exclude).";
			return args;
		}
	}

	if (args.captureMode == 1 || args.captureMode == 2) {
		if (params.find(L"pid") == params.end()) {
			args.errorMessage = L"Error: Missing required argument --pid for mode 1 or 2.";
			return args;
		}
		args.processId = std::wcstoul(params[L"pid"].c_str(), &endPtr, 10);
		if (endPtr == params[L"pid"].c_str() || *endPtr != L'\0' || args.processId == 0) {
			args.errorMessage = L"Error: Invalid process ID: " + params[L"pid"] +
				L"\nMust be a positive integer.";
			return args;
		}
	}

	args.isValid = true;
	return args;
}

void usage() {
	std::wcout << L"Process Audio Recorder - Captures audio from processes or the entire system\n\n"
		<< L"Usage: ProcessAudioRecorder [--pid <PID>] --mode <MODE> --path <FILEPATH>\n\n"
		<< L"Options:\n"
		<< L"  --pid <PID>    Target process ID (required for mode 1 and 2)\n"
		<< L"  --mode <MODE>  Capture mode (required):\n"
		<< L"                 0 - Global mode: Capture all system audio (system mix)\n"
		<< L"                 1 - Include mode: Capture target process and its children (default)\n"
		<< L"                 2 - Exclude mode: Capture all except target process and its children\n"
		<< L"  --path <PATH>  Output file path (required)\n\n"
		<< L"Examples:\n"
		<< L"  ProcessAudioRecorder --mode 0 --path C:\\system_audio.wav\n"
		<< L"  ProcessAudioRecorder --pid 1234 --mode 1 --path C:\\record.wav\n"
		<< L"  ProcessAudioRecorder --pid 5678 --path D:\\audio.wav\n\n"
		<< L"Exit conditions:\n"
		<< L"  - Target process exits (for mode 1 and 2)\n"
		<< L"  - User presses Ctrl+C\n";
}

bool IsProcessRunning(DWORD processId) {
	HANDLE process = OpenProcess(SYNCHRONIZE, FALSE, processId);
	if (process == NULL) {
		return false;
	}
	DWORD result = WaitForSingleObject(process, 0);
	CloseHandle(process);
	return (result == WAIT_TIMEOUT);
}

BOOL WINAPI ConsoleCtrlHandler(DWORD dwCtrlType) {
	if (dwCtrlType == CTRL_C_EVENT) {
		g_bStopCapture = true;
		return TRUE;
	}
	return FALSE;
}

void DisplayProgress(const std::chrono::seconds& duration) {
	auto totalSeconds = duration.count();
	auto hours = totalSeconds / 3600;
	auto minutes = (totalSeconds % 3600) / 60;
	auto seconds = totalSeconds % 60;

	std::wcout << L"\r● Recording [";
	std::wcout << std::setw(2) << std::setfill(L'0') << hours << L":"
		<< std::setw(2) << std::setfill(L'0') << minutes << L":"
		<< std::setw(2) << std::setfill(L'0') << seconds << L"]"
		<< L" - Press Ctrl+C to stop        "
		<< std::flush;
}

int wmain(int argc, wchar_t* argv[]) {
	_wsetlocale(LC_ALL, L"");
	SetConsoleTitleW(L"Cirong Process Audio Recorder");
	CommandLineArgs args = ParseCommandLine(argc, argv);
	if (!args.isValid) {
		if (!args.errorMessage.empty()) {
			std::wcout << args.errorMessage << L"\n\n";
		}
		usage();
		return 1;
	}
	CLoopbackCapture loopbackCapture;
	g_pCurrentCapture = &loopbackCapture;
	HRESULT hr;
	if (args.captureMode == 0) {
		std::wcout << L"Starting global audio capture (system mix)..." << std::endl;
		hr = loopbackCapture.StartGlobalCaptureAsync(args.outputPath.c_str());
	}
	else {
		if (!IsProcessRunning(args.processId)) {
			std::wcout << L"Error: Process with ID " << args.processId << L" does not exist or cannot be accessed." << std::endl;
			return 3;
		}
		bool includeProcessTree = (args.captureMode == 1);
		std::wcout << L"Starting audio capture for process " << args.processId
			<< L" (" << (includeProcessTree ? L"include" : L"exclude")
			<< L" mode)..." << std::endl;
		hr = loopbackCapture.StartCaptureAsync(args.processId, includeProcessTree, args.outputPath.c_str());
	}
	if (FAILED(hr)) {
		wil::unique_hlocal_string message;
		FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_ALLOCATE_BUFFER,
			nullptr, hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
			(PWSTR)&message, 0, nullptr);
		std::wcout << L"Failed to start capture\n0x" << std::hex << hr << L": " << message.get() << std::endl;
		return 2;
	}
	if (!SetConsoleCtrlHandler(ConsoleCtrlHandler, TRUE)) {
		std::wcout << L"Warning: Could not set control handler. Ctrl+C may not work properly." << std::endl;
	}
	std::wcout << L"Capture started successfully." << std::endl;
	std::wcout << L"Press Ctrl+C to stop capture at any time." << std::endl;
	auto startTime = std::chrono::steady_clock::now();
	bool processExited = false;
	bool userInterrupted = false;
	while (true) {
		if (g_bStopCapture) {
			userInterrupted = true;
			break;
		}
		if (args.captureMode != 0) {
			if (!IsProcessRunning(args.processId)) {
				processExited = true;
				break;
			}
		}
		auto currentTime = std::chrono::steady_clock::now();
		auto duration = std::chrono::duration_cast<std::chrono::seconds>(currentTime - startTime);
		DisplayProgress(duration);
		std::this_thread::sleep_for(std::chrono::milliseconds(500));
	}
	std::wcout << std::endl;
	if (userInterrupted) {
		std::wcout << L"Capture interrupted by user. Stopping..." << std::endl;
	}
	else if (processExited) {
		std::wcout << L"Target process has exited. Stopping capture..." << std::endl;
	}
	loopbackCapture.StopCaptureAsync();
	std::wcout << L"Finishing capture and saving file..." << std::endl;
	HANDLE hStopEvent = loopbackCapture.GetStopEventHandle();
	if (hStopEvent != NULL) {
	}
	std::wcout << L"Capture completed. File saved to: " << args.outputPath << std::endl;
	g_pCurrentCapture = nullptr;
	return 0;
}
