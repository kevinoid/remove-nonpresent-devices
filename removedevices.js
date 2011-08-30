/* Script to remove unnecessary (nonpresent) devices using DevCon
 *
 * WARNING:
 *   This script will remove/uninstall devices from your computer system.
 * Although several important types of devices have been excluded from the
 * default removal process, the possibility of removing devices critical to
 * the operation of your computer system still exists.  PLEASE double-check
 * the list of devices which will be removed before pressing "Y" and PLEASE
 * create a system backup if you are removely concerned about making your
 * system unbootable.
 *
 * Dependencies:
 * - devcon.exe must be in the same directory as this script
 *   (or in the search PATH)
 *   This file can be downloaded from Microsoft and is described at
 *   http://support.microsoft.com/kb/311272
 *
 * Usage:
 * - To remove all non-present devices (except legacy and software devices)
 *   either double-click the script or invoke it from the command prompt
 *
 * - To delete devices without confirmation, run
 *	removedevices.js /noconfirm
 *
 * - To see the output of DevCon as the script executes, run
 *	removedevices.js /verbose
 *
 * - To create a list of devices which would be removed, run
 *	removedevices.js /outfile:devicelist.txt
 *
 * - To delete all devices (device IDs) listed in a file, run
 *	removedevices.js /infile:devicelist.txt
 *
 * Changelog:
 *  2010-02-21	Fix argument passing when invoked with wscript
 *  2010-02-06  Initial release
 *
 * License:
 * Copyright (c) 2009, Kevin Locke <klocke@digitalenginesoftware.com>
 * 
 * Permission to use, copy, modify, and/or distribute this software for any
 * purpose with or without fee is hereby granted, provided that the above
 * copyright notice and this permission notice appear in all copies.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
 * MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
 * ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
 * WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 * ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 * OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 */

var WShell = WScript.CreateObject("WScript.Shell");

var options = {
    deleteall: { desc: "Delete all non-present devices (Dangerous)" },
    deletelegacy: { desc: "Also delete legacy non-present devices (Dangerous)" },
    deletesw: { desc: "Also delete software non-present devices (Dangerous)" },
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
    var swre = /^SW(?:MUXBUS)?\\/;
    nonpresentdevices = nonpresentdevices.filter(function(dev) {
	    // Note:  Boolean inverted for efficiency (only regex when needed)
	    return !((!deletelegacy && legacyre.test(dev.devid)) ||
		    (!deletesw && swre.test(dev.devid)) ||
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
