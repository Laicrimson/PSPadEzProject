/**-----------------------------------------------------------------------------
 * Filename		: EzProject.js
 * Last Modified	: 08/02/2010 10:24:31
 * Description		: This script refreshs the active project.
 * Created 		: 2 September 2009
 * Created by 		: PoYang Lai ( Lai.crimson@gmail.com )
 * Tested with		: PSPad 2383 and above
-----------------------------------------------------------------------------**/

/**-----------------------------------------------------------------------------
 * Description		: EzProject sync the PSPad virtual project files with true
 *                              directory specified in project default directory.
 *  -----------------------------------------------------------------------------**/

/**-----------------------------------------------------------------------------
 * Requirement		:
 *  -----------------------------------------------------------------------------**/

/**-----------------------------------------------------------------------------
 * Note			: It will set the project file (.ppr) to read only to
 *                              prevent the auto saving of PSPad when PSPad closed.
 *                        Use EzSaveProject to save project if you want to save
 *                              the read only project.
-----------------------------------------------------------------------------**/

/**-----------------------------------------------------------------------------
 * Version History	:
 * 0.130                : 08/09/2011
 *                        Support to skip the directory listed in SkipList.txt
 * 0.120                : 08/02/2010
 *                        1. Keep relative path feature if AbsolutePath = 0.
 *                        2. Fix bug if default dir is root of partition.
 * 0.110                : 13/01/2010
 *                        Treat all lines as config parameters before [Project tree]
 *                        tag and save all of them.
 * 0.100		: 02/09/2009
-----------------------------------------------------------------------------**/

var module_name = "EzProject";
var module_version = "0.120";

/**
 * Shortcut setting.
 * user can modify shortcut here.
 **/
//var gShortcutCreateProject = "";
var gShortcutRefreshProject = "ALT+F5";
var gShortcutSaveProject = "";

/**
 * Output strings.
 * All warning strings are listed here.
 **/
var sNoProject = "\nNo Project...\n";
var sProjFileNotExist = "\nProject File is not existed...\n";
var sNoDefaultDir = "\nNo default directory...\n";
var sRefreshProjectDone = "\nRefresh project done...\n";

/**-----------------------------------------------------------------------------
 * Programs
-----------------------------------------------------------------------------**/
//gloabl variables
var AbsolutePathID = "AbsolutePath";
var DefaultDirID = "DefaultDir";
var ProjectTreeTag = "[Project tree]";
var OpenProjectFilesTag = "[Open project files]";
var SelectedProjectFilesTag = "[Selected Project Files]";
// 1.20 >>>
var gAbsolutePathSetting = 1;
var gDefaultDir = "";
var gRelativePathPattern = "";
// 1.20 <<<
// 1.30 >>>
var gSkipList = new Array();
var SkipListFileName = "SkipList.txt";
var sInputPromptSkipList = "\nSkip Directory List in SkipList.txt    (Y/n)...\n";
// 1.30 <<<

var CmdShell = CreateObject("WScript.Shell");
var FSObj = CreateObject("Scripting.FileSystemObject");

function EzRefreshProject()
{
        var ProjFilesCnt = projectFilesCount();
        var ProjName, BackupName;
// 1.20        var DefaultDir = "";
        var TextStream, TextStream2;

	var i = 0;

	if(ProjFilesCnt<=0){
		echo(sNoProject);
		return;
	}

	// get project filename
        ProjName = projectFileName();
        if(ProjName == null || ProjName == ""){
	        echo(sProjFileNotExist);
                return;
        }

	BackupName = ProjName.concat(".bak");

	TextStream2 = FSObj.OpenTextFile(BackupName, 2, true, -2);
	
	// OpenTextFile(filename,
	//		1:ForReading 2:ForWriting 8:ForAppending,
	//		create,
	//		-2:TristateUseDefautl -1:TristateTrue 0:TristateFalse
	//		)
	TextStream = FSObj.OpenTextFile(ProjName, 1, true, -2);

	// Save config parameters between [Config] and [Project tree]
	for(i=0; !TextStream.AtEndOfStream; i++){
                // copy the configuration
		Lines = TextStream.ReadLine();

		if(Lines == ProjectTreeTag)
		        break;

// 1.20		// set absolute path = 1
// 1.20		if(Lines.slice(0, AbsolutePathID.length) == AbsolutePathID){
// 1.20		        Lines = AbsolutePathID.concat("=1");
// 1.20		}
// 1.20 >>>
		// get absolute path setting
                if(Lines.slice(0, AbsolutePathID.length) == AbsolutePathID){
		        var temp_lines = Lines.split("=");
                        gAbsolutePathSetting = parseInt(temp_lines[1]);
		}
// 1.20 <<<
	        // get the default dir
		if(Lines.slice(0, DefaultDirID.length) == DefaultDirID){
		        var tempstr = Lines.split("=");
                        gDefaultDir = String(tempstr[1]);
		}

	        TextStream2.WriteLine(Lines);
	}
	TextStream.Close();

	if(gDefaultDir == ""){
	        echo(sNoDefaultDir);
	        TextStream2.Close();
	        if(FSObj.FileExists(BackupName)){
	                FSObj.DeleteFile(BackupName);
	        }
	        return;
	}

// 1.20 >>>
	// if relative path
	if(gAbsolutePathSetting == 0){
		var lastslash = ProjName.lastIndexOf("\\");
		GetRelativePathPattern(ProjName.slice(0, lastslash), gDefaultDir);
	}
// 1.20 <<<

	TextStream2.WriteLine(ProjectTreeTag);
// 1.20 >>>
	// take care of root of partition
	if(gDefaultDir.length == 3){
		TextStream2.WriteLine(gDefaultDir);
	}else{
// 1.20 <<<
		var lastslash = gDefaultDir.lastIndexOf("\\");
		TextStream2.WriteLine(gDefaultDir.substring(lastslash+1));
// 1.20 >>>
	}
// 1.20 <<<

// 1.30 >>>
	// get the SkipList for EzProject module
	GetEzProjectSkipList();
// 1.30 <<<

	GetFileList(gDefaultDir, 1, TextStream2);
	TextStream2.WriteLine(OpenProjectFilesTag);
	TextStream2.WriteLine(SelectedProjectFilesTag);
	TextStream2.Close();

	var FileObj = FSObj.GetFile(ProjName);
	FileObj.Attributes &= (~1);     // clear read only
	FSObj.CopyFile(BackupName, ProjName, true);
	FileObj.Attributes |= 1;        // set read only
	FSObj.DeleteFile(BackupName);

	// record this action for EzProject module
	var EzProjectRecord = GetOnlyFilePath(moduleFileName(module_name));
	EzProjectRecord = EzProjectRecord.concat("\\EzProject\\record.txt");

	// OpenTextFile(filename,
	//		1:ForReading 2:ForWriting 8:ForAppending,
	//		create,
	//		-2:TristateUseDefautl -1:TristateTrue 0:TristateFalse
	//		)
	var TextStream3 = FSObj.OpenTextFile(EzProjectRecord, 8, true, -2);
	TextStream3.WriteLine(ProjName);
	TextStream3.Close();

	// finish refresh project
	echo(sRefreshProjectDone);

// 	var PspadExe = getVarValue("%PSPath%");
// 	PspadExe = PspadExe.concat("\\PSPad.exe");
// 	CmdShell.Run(PspadExe + " " + ProjName, 0, true);
	return;
}

// 1.30 >>>
function GetEzProjectSkipList()
{
	var EzProjectSkipList = GetOnlyFilePath(moduleFileName(module_name));
	EzProjectSkipList = EzProjectSkipList.concat("\\EzProject\\" + SkipListFileName);
	var TextStream = FSObj.OpenTextFile(EzProjectSkipList, 1, true, -2);
	var lines, lastslash;

	var yes_no = inputText( sInputPromptSkipList, null, null);
	if(yes_no != "y" && yes_no != "Y") return;

	for(var i=0, j=0; !TextStream.AtEndOfStream; i++){
		lines = TextStream.ReadLine();
	  	lastslash = lines.lastIndexOf("\\");
		if(lastslash == (lines.length-1)){
			lines = lines.substring(0,lastslash);
		}
  		gSkipList[j++] = gDefaultDir + "\\" + lines;
	}
}
// 1.30 <<<

function GetRelativePathPattern(projpath, defaultpath)
{
	var pp_array = projpath.split("\\");
	var dp_array = defaultpath.split("\\");
	var i, j;
	for(i=0;i<pp_array.length && i<dp_array.length;i++){
		if(pp_array[i] != dp_array[i])
			break;
	}
	if(i<pp_array.length){
		for(j=i;j<pp_array.length;j++){
			if(gRelativePathPattern.length == 0)
				gRelativePathPattern = "..";
			else
				gRelativePathPattern += "\\..";
		}
	}
	if(i<dp_array.length){
		for(j=i;j<dp_array.length;j++){
			gRelativePathPattern += "\\";
			gRelativePathPattern += dp_array[j];
		}
	}
}

function GetOnlyFileName(file_str)
{
        var lastslash = file_str.lastIndexOf("\\");
        return file_str.substring(lastslash+1);
}

function GetOnlyFilePath(file_str)
{
        var lastslash = file_str.lastIndexOf("\\");
        return file_str.substring(0,lastslash);
}

function GetFileList(dir, depth, outputfile)
{
	var fs_dir, fs_subdir, fs_file, fd;
        var tab = "";
        for(var i=0;i<depth;i++) tab += "\t";

// 1.30 >>>
	for(var i=0; i < gSkipList.length; i++){
		if(dir == gSkipList[i]) return;
	}
// 1.30 <<<

	fs_dir = FSObj.GetFolder(dir);
	fs_subdir = new Enumerator(fs_dir.SubFolders);
	for (; !fs_subdir.atEnd(); fs_subdir.moveNext()){
	        outputfile.WriteLine(tab+"-"+GetOnlyFileName(String(fs_subdir.item())));
	        GetFileList(fs_subdir.item(), depth+1, outputfile);
	}
	fs_file = new Enumerator(fs_dir.Files);
	for (; !fs_file.atEnd(); fs_file.moveNext()){
// 1.20 >>>
		// if AbsolutePath=1
		if(gAbsolutePathSetting){
// 1.20 <<<
			outputfile.WriteLine(tab+fs_file.item());
// 1.20 >>>
		// if AbsolutePath=0
		}else{
			var relative_file = String(fs_file.item());
			// if gRelativePathPattern is NULL string, don't show '\\' at first character
			if(gRelativePathPattern.length == 0)
				outputfile.WriteLine(tab+gRelativePathPattern+relative_file.substring(gDefaultDir.length+1));
			else
				outputfile.WriteLine(tab+gRelativePathPattern+relative_file.substring(gDefaultDir.length));
		}
// 1.20 <<<
	}
}

function EzSaveProject()
{
        var ProjName;

	// get project filename
        ProjName = projectFileName();
        if(ProjName == null || ProjName == ""){
	        echo(sProjFileNotExist);
                return;
        }
        
	var FileObj = FSObj.GetFile(ProjName);
	FileObj.Attributes &= (~1);     // clear read only
	runPSPadAction("aProjSave");
}

function OpenActivateFile( file_name)
{
        var NewEditObj = newEditor(); //New editor object
	if(FSObj.FileExists(file_name)){
		NewEditObj.openFile(file_name);
	}else{
	        echo(sFileNotExist);
	        return;
	}
	NewEditObj.activate();
	return;
}

function OpenModule()
{
	try{
                OpenActivateFile( moduleFileName(module_name));
	}
	catch(e){
		echo("\nOpen file error...'\n" + moduleFileName(module_name) + "\n" + e.message + "\n");
	}
	return;
}

function Init()
{
	var EzProjectRecord = GetOnlyFilePath(moduleFileName(module_name));
	EzProjectRecord = EzProjectRecord.concat("\\EzProject\\record.txt");

	// OpenTextFile(filename,
	//		1:ForReading 2:ForWriting 8:ForAppending,
	//		create,
	//		-2:TristateUseDefautl -1:TristateTrue 0:TristateFalse
	//		)
	if(FSObj.FileExists(EzProjectRecord)){
		var TextStream = FSObj.OpenTextFile(EzProjectRecord, 1, true, -2);
		while(!TextStream.AtEndOfStream){
			var FileName = TextStream.ReadLine();
			if(FSObj.FileExists(FileName)){
				var FileObj = FSObj.GetFile(FileName);
				FileObj.Attributes &= (~1);
			}
		}
		TextStream.Close();
		FSObj.DeleteFile(EzProjectRecord);
	}

	addMenuItem("Ez&RefreshProject", "EzProject", "EzRefreshProject", gShortcutRefreshProject);
	addMenuItem("Ez&SaveProject", "EzProject", "EzSaveProject", gShortcutSaveProject);
	addMenuItem("-", "EzProject", "", "");
	addMenuItem("&EditEzProject", "EzProject", "OpenModule", "");
}
