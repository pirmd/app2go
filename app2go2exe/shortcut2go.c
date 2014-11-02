/**
   Copyright (c) 2014, Pierre Duranton <pmd-prj@laposte.net>
   All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions
  are met:
  1. Redistributions of source code must retain the above copyright
     notice, this list of conditions and the following disclaimer.
  2. Redistributions in binary form must reproduce the above copyright
     notice, this list of conditions and the following disclaimer in the
     documentation and/or other materials provided with the distribution.

  THIS SOFTWARE IS PROVIDED BY AUTHOR AND CONTRIBUTORS ``AS IS'' AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
  ARE DISCLAIMED.  IN NO EVENT SHALL AUTHOR OR CONTRIBUTORS BE LIABLE
  FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
  DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS
  OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
  HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
  LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY
  OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF
  SUCH DAMAGE.
**/

/**
 TODO:
 [ ] Instead of FreeConsole(), consider:
  .http://stackoverflow.com/questions/6342935/start-a-program-without-a-console-window-in-background
  .http://stackoverflow.com/questions/8945018/hiding-the-black-window-in-c
 [ ] Improve command line quoting:
  .http://blogs.msdn.com/b/twistylittlepassagesallalike/archive/2011/04/23/everyone-quotes-arguments-the-wrong-way.aspx?Redirected=true
 [ ] use secure version of functions
**/

#include <stdlib.h>
#include <string.h>
#include <stdio.h>
#include <errno.h>

const char *APP2GOPATH = "APP2GOPATH=.;..;.\\App2GoPlatform;..\\App2GoPlatform;..\\..\\App2GoPlatform;Tools\\App2GoPlatform;Outils\\App2GoPlatform;bin\\App2GoPlatform;app2go\\App2GoPlatform";

void str_remove(char *s1, const char *s2) {
    char *p = s1;
    do {
        if ((p = strstr(p, s2)) != NULL) {
            memmove(p, p + strlen(s2), strchr(p, '\0') - p);
        }
    } while (p != NULL);
}


int main(int argc, char *argv[]) {
   FreeConsole();

   // Find useful information about my script
   char my_dir[_MAX_DIR];
   char my_name[_MAX_FNAME];
   char my_drive[_MAX_DRIVE];
   char my_folder[_MAX_PATH];
   _splitpath(argv[0], my_drive, my_dir, my_name, NULL);
   _makepath(my_folder, my_drive, my_dir, NULL, NULL);

   // Find app2go script within possible locations
   char working_path[_MAX_PATH];
   _getcwd(working_path, _MAX_PATH);

   _chdir(my_folder);
   if ( getenv("APP2GOPATH") == NULL ) {
        _putenv(APP2GOPATH);
   }
   char app2go_launcher[_MAX_PATH];
   _searchenv("app2go.wsf", "APP2GOPATH", app2go_launcher);
   if( *app2go_launcher == '\0' ) {
       AllocConsole();
       fprintf(stderr, "Shortcut is broken!\n App2go.wsf cannot be found from %s looking for %s.\n", my_folder, APP2GOPATH);
       system("pause");
       exit(EXIT_FAILURE);
    }
    _fullpath(app2go_launcher, app2go_launcher, _MAX_PATH);
    _chdir(working_path);
    //printf("app2go.wsf in %s\n", app2go_launcher);

   // Guess script name from programm name
   char *app2go_script = my_name;
   str_remove(app2go_script, "Portable");
   str_remove(app2go_script, "2go");
   //printf("script to run is %s\n", app2go_script);

   // Build command line to execute and run it
   char *to_exec[argc+4];
   to_exec[0] = "wscript.exe";
   to_exec[1] = "/NOLOGO";
   to_exec[2] = app2go_launcher;
   to_exec[3] = app2go_script;
   AllocConsole();
   int count = 0;
   int size = 0;
   for(count = 0; count < argc-1; count++){
         if ( strstr(argv[count+1], " ") != NULL ) {
             size = strlen(argv[count+1]) + 3;
             to_exec[count+4] = malloc(size);
             snprintf(to_exec[count+4], size, "\"%s\"", argv[count+1]);
         }
         else {
             to_exec[count+4] = argv[count+1];
         }
   }
   to_exec[count+4] = NULL;

   _execvp(to_exec[0], to_exec);
   AllocConsole();
   perror("execlp");
   exit(EXIT_FAILURE);
}
