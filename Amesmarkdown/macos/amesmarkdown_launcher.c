#include <fcntl.h>
#include <stdlib.h>
#include <unistd.h>

extern char **environ;

int main(void) {
    const char *app_root = "/Users/pensong/Documents/Amesmarkdown";
    const char *python = "/Users/pensong/Documents/Amesmarkdown/.venv/bin/python";
    const char *log_path = "/Users/pensong/Documents/Amesmarkdown/macos/amesmarkdown-app.log";

    setenv("PYTHONUNBUFFERED", "1", 1);
    setenv("PYTHONPATH", "/Users/pensong/Documents/Amesmarkdown/src", 1);
    chdir(app_root);

    int fd = open(log_path, O_CREAT | O_WRONLY | O_APPEND, 0644);
    if (fd >= 0) {
        dup2(fd, STDOUT_FILENO);
        dup2(fd, STDERR_FILENO);
        close(fd);
    }

    char *const argv[] = {"Amesmarkdown", "-m", "mdconvert_app.gui", NULL};
    execve(python, argv, environ);
    return 127;
}
