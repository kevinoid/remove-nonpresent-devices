/*
 * Remove Non-present Devices using DevCon
 * =======================================
 * 
 * removedevices.js is a script to automatically remove all non-present (i.e.
 * disconnected) devices from a Windows computer system.  This can often be
 * useful to prevent misbehaving and/or unnecessary drivers from being loaded.
 * 
 * WARNING
 * -------
 * This script will remove/uninstall devices from your computer system.
 * Although several important types of devices have been excluded from the
 * default removal process, the possibility of removing devices critical to
 * the operation of your computer system still exists.  PLEASE double-check
 * the list of devices which will be removed before pressing "Y" and PLEASE
 * create a system backup if you are remotely concerned about making your
 * system unbootable.
 * 
 * Examples
 * --------
 *  - To remove all non-present devices (except legacy and software devices)
 *    either double-click the script or invoke it from the command prompt
 *    without options.
 * 
 *  - To delete devices without confirmation, run
 * 
 *         removedevices.js /noconfirm
 * 
 *  - To see the output of DevCon as the script executes, run
 * 
 *         removedevices.js /verbose
 * 
 *  - To create a list of devices which would be removed, run
 * 
 *         removedevices.js /outfile:devicelist.txt
 * 
 *  - To delete all devices (device IDs) listed in a file, run
 * 
 *         removedevices.js /infile:devicelist.txt
 * 
 * 
 * Installation Instructions
 * =========================
 * 
 * This script requires that devcon.exe be available, either in the same
 * directory as this script or in a directory included in the PATH environment
 * variable.
 * 
 * For Windows XP and earlier, devcon.exe can be downloaded from Microsoft as
 * described in <http://support.microsoft.com/kb/311272>.
 * 
 * For Windows Vista and later, a newer version of devcon.exe is required.
 * It can be extracted from the Windows WDK as described in 
 * <http://social.technet.microsoft.com/wiki/contents/articles/how-to-obtain-the-current-version-of-device-console-utility-devcon-exe.aspx>.
 * 
 * 
 * ChangeLog
 * =========
 * 
 * 2012-08-02  Ryan Pavlik <abiryan@ryand.net>
 *
 *     * removedevices.js: Default to not removing extra volume shadow copies
 *
 * 2011-08-30  Kevin Locke <klocke@digitalenginesoftware.com>
 * 
 *     * *.*: Add typical package accoutrements (README, ChangeLog, etc)
 * 
 * 2011-02-10  Kevin Locke <klocke@digitalenginesoftware.com>
 * 
 *     * removedevices.js: Fix crash when relaunching with CScript
 *     * removedevices.js: Add note about devcon version for Vista and later
 * 
 * 2010-02-21  Kevin Locke <klocke@digitalenginesoftware.com>
 * 
 *     * removedevices.js: Initial Release
 * 
 * 
 * COPYING
 * =======
 *
 * Copyright 2010-2011 Digital Engine Software, LLC
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

var WShell = WScript.CreateObject("WScript.Shell");

var options = {
    deleteall: { desc: "Delete all non-present devices (Dangerous)" },
    deletelegacy: { desc: "Also delete legacy non-present devices (Dangerous)" },
    deletesw: { desc: "Also delete software non-present devices (Dangerous)" },
    deletevss: { desc: "Also delete volume shadow service non-present devices (potentially dangerous)" },
    infile: { desc: "Remove devices listed in a file (same format as devcon output)",
	      has_arg: true },
    help: { desc: "Print this help message" },
    noconfirm: { desc: "Do not ask for confirmation before deleting devices" },
    outfile: { desc: "Output list of devices which would be removed",
	       has_arg: true },
    quiet: { desc: "Do not print any unnecessary messages" },
    verbose: { desc: "Print extra diagnostic information as the program executes" },
    "?": { desc: "Print this help message" }
};

// Make sure we are running with CScript so user can see output
if (!/cscript\.exe$/i.test(WScript.FullName)) {
    var cmdline = "cscript.exe";
    var args = WScript.Arguments;
    cmdline += " \"" + WScript.ScriptFullName + "\"";
    for (var i=0; i<args.length; ++i)
	cmdline += " \"" + args(i) + "\"";

    WShell.Run(cmdline, 1, false);
    WScript.Quit(0);
}

// Array.filter (from ECMA-262 Edition 5)
// Code from Mozilla Developer Center
if (!Array.prototype.filter) {
    Array.prototype.filter = function(fun /*, thisp*/) {
        var len = this.length >>> 0;
        if (typeof fun != "function")
            throw new TypeError();

        var res = [];
        var thisp = arguments[1];
        for (var i = 0; i < len; i++) {
            if (i in this) {
                var val = this[i]; // in case fun mutates this
                if (fun.call(thisp, val, i, this))
                    res.push(val);
            }
        }

        return res;
    };
}

function assert(cond /*, msg */) {
    if (cond)
	return;

    var msg = arguments[1];
    if (msg == null)
	msg = "Assertion Failed";

    WScript.Echo("Internal Program Error: " + msg);
    WScript.Quit(1);
}

function combine_values(v1, v2) {
    if (v1 == null) {
	return v2;
    } else if (v1 instanceof Array) {
	v1.push(v2);
	return v1;
    } else {
	return new Array(v1, v2);
    }
}

function get_options(options, args) {
    var optvals = new Array();

    var optregex = /^(?:\/|--)([^:]+)(?:(?::|=)(.+))?/;
    for (var i=0; i<args.length; ++i) {
	var arg = args(i);	// Careful of the syntax... JScript...

	var argmatch = optregex.exec(arg);
	if (argmatch == null) {
	    optvals.push(arg);
	    continue;
	}

	argname = argmatch[1];
	argval = argmatch[2];
	if (argname in options) {
	    var option = options[argname];
	    var optval = optvals[argname];
	    if (option.has_arg && (argval || i<args.length-1)) {
		if (argval)
		    optvals[argname] = combine_values(optval, argval);
		else
		    // Careful again...
		    optvals[argname] = combine_values(optval, args(++i));
	    } else
		optvals[argname] = combine_values(optval, true);
	} else {
	    throw new Error("Unrecognized option: " + arg);
	}
    }

    return optvals;
}

function get_nonpresent_devices() {
    if (optvals["infile"])
	try {
	    return read_devids(optvals["infile"]);
	} catch (ex) {
	    WScript.Echo("Error:  Unable to read input file \"" +
		    optvals["infile"] + "\":  " + ex.message);
	    WScript.Quit(1);
	}

    var deletelegacy = !!optvals["deletelegacy"];
    var deletesw = !!optvals["deletesw"];
    if (optvals["deleteall"])
	deletelegacy = deletesw = true;

    var alloutput, presentoutput;
    try {
	alloutput = run_program("devcon.exe findall *").output
	presentoutput = run_program("devcon.exe find *").output;
    } catch (ex) {
	WScript.Echo("Error:  Unable to execute devcon.exe:  " + ex.message);
	WScript.Quit(1);
    }

    var alldevices, presentdevices;
    try {
       alldevices = parse_devcon_output(alloutput);
       presentdevices = parse_devcon_output(presentoutput);
    } catch (ex) {
	WScript.Echo("Error:  Unable to parse devcon output:  " + ex.message);
	WScript.Quit(1);
    }

    var nonpresentdevices = new Array(alldevices.length - presentdevices.length);
    var ai = 0, ni = 0;
    for (var pi=0; pi<presentdevices.length; ++ai, ++pi) {
	var presentdevice = presentdevices[pi];

	while (alldevices[ai].devid < presentdevice.devid)
	    nonpresentdevices[ni++] = alldevices[ai++];
	assert(alldevices[ai].devid == presentdevice.devid,
		"Device listed in find but not in findall!?");
    }
    while (ai < alldevices.length)
	nonpresentdevices[ni++] = alldevices[ai++];
    assert(ni == nonpresentdevices.length, "All devices not accounted for");

    var htreere = /^HTREE\\/;
    var legacyre = /^ROOT\\LEGACY_/;
    var vssre = /^STORAGE\\VOLUMESNAPSHOT/;
    var swre = /^SW(?:MUXBUS)?\\/;
    nonpresentdevices = nonpresentdevices.filter(function(dev) {
	    // Note:  Boolean inverted for efficiency (only regex when needed)
	    return !((!deletelegacy && legacyre.test(dev.devid)) ||
		    (!deletesw && swre.test(dev.devid)) ||
			(!deletevss && vssre.test(dev.devid)) ||
		    htreere.test(dev.devid));
	});

    return nonpresentdevices;
}

function parse_devcon_output(output) {
    // Note:  Restrictions loosened so we can use it for reading input files
    var devregex = /^(\S+)\s*(?::\s*(.*))?$/gm;
    var devices = new Array();

    var match;
    while ((match = devregex.exec(output)) != null)
	devices.push({devid: match[1], devname: match[2]});

    devices.sort(function(a, b) {
	    if (a.devid < b.devid)
		return -1;
	    else if (b.devid < a.devid)
		return 1;
	    else
		return 0;
	});

    return devices;
}

function print_help() {
    WScript.StdOut.WriteLine("Device Removal Script\n" +
	    "Supported Options:");
    for (var optname in options) {
	var optnamelen = optname.length;
	var padsize = 19 - optnamelen;
	var padding = new Array(padsize > 0 ? padsize : 0).join(' ');
	WScript.StdOut.WriteLine('/' + optname + padding + options[optname].desc);
    }
}

function read_devids(filename) {
    var file;
    if (filename == "-")
	file = WScript.StdIn;
    else {
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	file = fso.OpenTextFile(filename, 1);
    }

    var devices = parse_devcon_output(file.ReadAll());

    if (filename != "-")
	file.Close();

    return devices;
}

// Note:  I am not aware of a safe way to read stdout and stderr separately
//        without risking a deadlock, unless the output can be predicted.
//        So we read the combined output here...
// See the discussion at http://www.tech-archive.net/Archive/Scripting/microsoft.public.scripting.wsh/2004-10/0287.html
// for more details
function run_program(cmd) {
    var wsexec = WShell.Exec(cmd + " 2>&1");
    wsexec.StdIn.Close();
    
    var output = "";
    while (wsexec.Status == 0)
	output += wsexec.StdOut.ReadAll();

    return { exitcode: wsexec.ExitCode, output: output };
}

function write_devids(filename, devices) {
    var file;
    if (filename == "-")
	file = WScript.StdOut;
    else {
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	file = fso.CreateTextFile(filename, true);
    }

    for (var i=0; i<devices.length; ++i)
	file.WriteLine(devices[i].devid);

    if (filename != "-")
	file.Close();
}

var optvals = get_options(options, WScript.Arguments);

if (("help" in optvals) || ("?" in optvals)) {
    print_help();
    WScript.Quit(0);
}

var quiet = !!optvals["quiet"];
var verbose = !!optvals["verbose"];

nonpresentdevices = get_nonpresent_devices();

if (optvals["outfile"]) {
    try {
	write_devids(optvals["outfile"], nonpresentdevices);
    } catch (ex) {
	WScript.Echo("Error:  Unable to write output file:  " + ex.message);
	WScript.Quit(1);
    }

    WScript.Quit(0);
}

if (!optvals["noconfirm"]) {
    WScript.StdOut.WriteLine("Devices which will be removed:");
    for (var i=0; i<nonpresentdevices.length; ++i) {
	var npd = nonpresentdevices[i];
	WScript.StdOut.WriteLine(npd.devid)
	if (npd.devname)
	    WScript.StdOut.WriteLine("\t" + npd.devname)
    }

    WScript.StdOut.Write("Are you sure (Y/N)? ");
    var answer = WScript.StdIn.Read(1);
    if (answer != "y" && answer != "Y")
	WScript.Quit(0);
}

for (var i=0; i<nonpresentdevices.length; ++i) {
    var device = nonpresentdevices[i];
    try {
	if (!quiet)
	    WScript.StdOut.Write(
		    "Removing " + device.devid + ":\n");

	var result = run_program("devcon.exe remove \"@" + device.devid + "\"");

	if (verbose) {
	    WScript.StdOut.WriteLine(result.output);
	    WScript.StdOut.WriteLine(
		    "Removal finished exit code " + result.exitcode);
	}
    } catch (ex) {
	WScript.StdOut.WriteLine("DevCon encountered an error: " + ex.message);
    }
}

if (!quiet)
    WScript.StdOut.WriteLine("All Done");
