#include <windows.h>
#include <stdio.h>

// Simple test app to run the dll outside of vim.
// Can be used to mail a properly formatted file from, say, a batch file.
// To build: cl mailit.cpp
// To use: mailit full_path_to_properly_formatted_file

int main(int argc, char* argv[]) {
	typedef char* (*mailer)(char*);

	if (argc != 2) {
		printf("Usage: mailit full_path_to_file\n");
		printf("The file must be formatted as described in vimmailr.vim\n");
		return 1;
	}

	int rc = 1;
	HINSTANCE hLib = LoadLibrary("vimmailr.dll");
	if (hLib) {
		mailer m = (mailer)GetProcAddress(hLib, "VimSendMail");
		if (m) {
			char* r = m(argv[1]);
			if (r) {
				printf("Return val: %s\n", r);
				if (!strcmp("Succeeded", r))
					rc = 0;
			}
		}
		else
			printf("Failed to find VimSendMail export\n");
		FreeLibrary(hLib);
	}
	else {
		printf("Failed to load library\n");
	}

	return rc;
}

