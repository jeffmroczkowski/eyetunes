//
// Script to convert your iTunes library to scale audio files
// to your start/end time edits
// 
// Copyright (C) 2014  Jeff Mroczkowski
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
// 
// Apple JavaScript COM program to scale all files *other than m4p files* 
// in place in your itunes collection if the files have had the start/end 
// times modified.  The original files are stored in the backup folder identified
// by the path in backupFolder below.  Please note that before you run this 
// program that you modify this variable and also set the converter options.
// See Preferences->General->Import Settings within iTunes. 
//
// USE THIS AT YOUR OWN RISK.  NO WARRANTY OR LIABILTY
//

var iTunesApp = WScript.CreateObject("iTunes.Application");
var mainLibrary = iTunesApp.LibraryPlaylist;
var mainLibrarySource = iTunesApp.LibrarySource;
var tracks = mainLibrary.Tracks;
var numTracks = tracks.Count;

var fso  = new ActiveXObject("Scripting.FileSystemObject"); 
var so = new ActiveXObject("WScript.Shell");
var fh = fso.CreateTextFile("log.txt", true); 

// change this to where you want to back files up to
var backupFolder = "C:\\Users\\Jeff\\Music\\backup\\";
// Just iterate and log stuff - no operations
var test = false;

main();

function main ()
{
    // create backupFolder
    if (!fso.FolderExists(backupFolder)) {
        WScript.Stdout.Write("Backup folder does not exist: creating: " + backupFolder);
        CreateDirs (backupFolder);   
    }

    // preserve existing date
    var dateNow = new Date();

    // Monitor for events
    WScript.ConnectObject(iTunesApp, "ITEventTest_");

    var i = 1;
    WScript.Stdout.Write("Processing tracks: " + numTracks);
    WScript.Stdout.Write("RUNNING TEST MODE " + test);
    for (i = 1; i <= numTracks; i++)
    {
        var currTrack = tracks.Item(i);
        var album = currTrack.Album;
        var artist = currTrack.Artist;
        var source = currTrack.Location;
        var song = currTrack.Song;
        var start = currTrack.Start;
        var finish = currTrack.Finish;
        var duration = currTrack.duration;
        var modified = currTrack.modificationDate;
        var size = currTrack.Size;
        var create = String(currTrack.DateAdded);
        var lists = currTrack.Playlists;

        // No real source location - can't do anything. - iCloud
        if (source == undefined) {
            fh.WriteLine ('Skipping missing source for : ' + i + ' Artist: ' + artist + ' Album: ' + album);
            continue;
        }

        fh.WriteLine ('For: ' + i + ' Artist: ' + artist + ' Album: ' + album +  ' source ' + source);

        var idx = source.lastIndexOf ('\\');
        var name = source.substr (idx + 1, source.length);
        var target = backupFolder + name;

        WScript.Stdout.Write('File ' + name + ' start ' + start + ' finish ' + finish + ' duration ' + duration + ' modified ' + modified + '\n');

        if (!fso.FileExists(source)) {
            var msg = 'Source file does not exist: exiting: this should not happen ' + source;
            WScript.Stdout.Write(msg);
            fh.WriteLine (msg);
        //    return;
        }


        // can't convert these old files.  Just leave them alone
        if (source.indexOf ('.m4p') > 0) {
            fh.WriteLine ('Skipping m4p : ' + source);
            WScript.Stdout.Write ('Skipping file cannot convert m4p : ' + source);
        }
        // if file has an edit
        else if (start != 0 || finish != duration) {

            //Sun Jan 20 20:33:16 PST 2013
            var pieces = create.split(" ");
            var monStr = pieces[1];
            if (monStr == "Jan") monStr = 1;
            if (monStr == "Feb") monStr = 2;
            if (monStr == "Mar") monStr = 3;
            if (monStr == "Apr") monStr = 4;
            if (monStr == "May") monStr = 5;
            if (monStr == "Jun") monStr = 6;
            if (monStr == "Jul") monStr = 7;
            if (monStr == "Aug") monStr = 8;
            if (monStr == "Sep") monStr = 9;
            if (monStr == "Oct") monStr = 10;
            if (monStr == "Nov") monStr = 11;
            if (monStr == "Dec") monStr = 12;

            // Lets change the date on the command line to preserve the create date time
            // single the create date attribute is read only.  We are going to reinsert 
            // the file using its old create date by going back in time.

            var dateCmd = "CMD.EXE  /C DATE ";
            dateCmd += monStr;
            dateCmd += "/";
            dateCmd += pieces[2];
            dateCmd += "/";
            dateCmd += pieces[5];
            //  var out = so.Exec ("CMD.EXE  /C DATE 12/22/2008");
            WScript.Stdout.Write ("Date of file : " + create + "\n");
            WScript.Stdout.Write ("Changing date " + dateCmd + "\n");
            fh.WriteLine ('Changing date ' + dateCmd);
            so.Exec (dateCmd);

            if (!fso.FileExists(target)) {
                if (!test) {
                    fso.CopyFile (source, target, true);
                    fh.WriteLine ('Creating backup of file: ' + source + ' to ' + target);
                }
            }
            else { // file exists check for size difference
                var f = fso.GetFile(target);
                if (f.Size != size) {
                    fh.WriteLine ('Not continuing since backup exists but different size: ' + target);
                    return;
                }
                WScript.Stdout.Write('backup already exists: ' + target + '\n');
            }

            // Convert to file
            WScript.Stdout.Write('Would convert file ' + i + " Name "  + source + '\n');

            try {

                if (!test) {
                    var newFile = ConvertThisTrack(source);
                    WScript.Stdout.Write('NEWFILE ' + newFile + '\n');
                    fh.WriteLine ('Count ' + i + ' Artist: ' + artist + ' Album: ' + album);
                    currTrack.Delete();  // delete original

                    WScript.Stdout.Write('Playlist Count ' + (lists.Count-1) + '\n');
                    for (var l = 1; l < lists.Count; l++) {

                        var playListItem = lists.Item(l);
                        if (playListItem.Name == "Music" || playListItem.Smart) {
                        //if (playListItem.Name == "Music" || playListItem.Name.indexOf ('#') == 0  || playListItem.Name.indexOf ('*') == 0) {
                            fh.WriteLine ('Not inserting into main collection or smart playlist ' + playListItem.Name);
                            WScript.Stdout.Write ('Not inserting into main collection or smart playlist ' + playListItem.Name + "\n");
                        }
                        else {
                            WScript.Stdout.Write ("Added file to playlist: " + playListItem.Name + "\n");
                            playListItem.AddFile(newFile);
                        }
                    }
                }
            }
            catch (e) {
                WScript.Stdout.Write ("Unable to convert " + source);
                return;
            }
        }
    }

    var nowCmd = "CMD.EXE  /C DATE ";
    nowCmd += dateNow.getMonth() + 1;
    nowCmd += "/";
    nowCmd += dateNow.getDay();
    nowCmd += "/";
    nowCmd += dateNow.getYear();
    WScript.Stdout.Write ("Changing date back to original " + nowCmd + "\n");
    so.Exec (nowCmd);
    fh.Close(); 
}

function ConvertThisTrack(track) 
{
    var opStatus = iTunesApp.ConvertFile2(track);
    var name = opStatus.trackName;
    WScript.Stdout.Write('Converting file ' + name + "\n");
    fh.WriteLine ('Converting file ' + name + '\n');
    while (opStatus.InProgress) 
    {
        WScript.Sleep (500);
        WScript.Stdout.Write('Sleeping ' + name + "\n");
    }
    if (opStatus.Tracks.Item(1).Location == undefined || opStatus.Tracks.Item(1).Location == null || opStatus.Tracks.Item(1).Location == "") {
        fh.WriteLine("throwing exception");
        throw Exception ("Failed to convert");
    }
    var outFile = fso.GetFile(opStatus.Tracks.Item(1).Location);
    WScript.Stdout.Write('Converted file: ' + outFile);
    return outFile;
}

function CreateDirs(dir)
{
    var pieces = dir.split("\\");
    var subdir = "";
    var i;
    for (i = 0; i < pieces.length; i++) {
        subdir += pieces[i];
        if (fso.FolderExists(subdir)) {
                fh.WriteLine ('Folder exists ' + subdir);
        }
        else {
                fh.WriteLine ('Try folder create ' + subdir);
                fso.CreateFolder(subdir)
                fh.WriteLine ('Folder create ' + subdir);
        }
        subdir += "\\";
    }
}

function ITEventTest_OnDatabaseChangedEvent(deletedObjects, changedObjects)
{
    return;
}
