trebuie sa ai `pip` instalat

daca `pip` iti da eroare, il adaugi la environment variables / user / PATH & restart cmd terminal.

Instalare:
```
pip install pyinstaller
```

Utilizare: #TODO - incercat care varianta e mai buna

I suspect that you're using pyinstaller's "one file" mode -- this mode means that it has to unpack all of the libraries to a temporary directory before the app can start. In the case of Qt, these libraries are quite large and take a few seconds to decompress. Try using the "one directory" mode and see if that helps?

https://coderslegacy.com/pyinstaller-spec-file-tutorial/

https://pyinstaller.org/en/v3.2/usage.html

https://www.devdungeon.com/content/pyinstaller-tutorial



