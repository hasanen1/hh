var code = 'x=msgbox("Hello" ,0, "test")';

function wrap(s, c) {
    return c + s + c;
}

var sc = null;
try {
    sc = new ActiveXObject("MSScriptControl.ScriptControl");
}
catch (err) {
    var wsc32path = "%SystemRoot%\\SysWOW64\\cscript.exe";
    var cmd = wrap(wsc32path, "\"") + " /e:jscript " + wrap(WScript.ScriptFullName, "\"");
    var shell = WScript.CreateObject("WScript.Shell");
    shell.Run(cmd, 0, false);
    WScript.Quit(0);
}

sc.Language = "VBScript";
sc.Timeout = -1;
sc.AllowUI = true;
sc.AddObject("wscript", WScript, true);
sc.AddCode(code);

