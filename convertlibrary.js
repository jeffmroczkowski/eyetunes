//
// Script to sync and convert your iTunes library to a remote drive.
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
// Apple JavaScript COM program to sync your iTunes library to a remote drive
// converting all files that don't exist on the remote drive to whatever your
// converter preferences are set to (Preferences->General->Import Settings) 
// within iTunes.  Before running check sourceDrive, remoteDrive and xmlPath
// to make sure it matches your environment.  If not change.  
//
// USE THIS AT YOUR OWN RISK.  NO WARRANTY OR LIABILTY
//

var iTunesApp = WScript.CreateObject("iTunes.Application");
var mainLibrary = iTunesApp.LibraryPlaylist;
var mainLibrarySource = iTunesApp.LibrarySource;
var tracks = mainLibrary.Tracks;
var numTracks = tracks.Count;

var fso  = new ActiveXObject("Scripting.FileSystemObject"); 
var so = new ActiveXObject("shell.application");
var fh = fso.CreateTextFile("log.txt", true); 

// CHANGE THESE FOR YOUR ENVIRONMENT
var sourceDrive = 'C';
var remoteDrive = 'Z';
var xmlPath = "\\Users\\Jeff\\Music\\iTunes\\iTunes Music Library.xml";
// CHANGE THESE FOR YOUR ENVIRONMENT

main();

function main ()
{
    if (!fso.DriveExists (remoteDrive + ':')) {
        WScript.Stdout.Write("Remote drive does not exist " + remoteDrive + "\n");
        return;
    }

    var i = 1;
    WScript.Stdout.Write("Processing tracks: " + numTracks);
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
        fh.WriteLine ('For: ' + i + ' Artist: ' + artist + ' Album: ' + album);
        var size = currTrack.Size;

        var target = source.substr (1, source.length);
        target = remoteDrive + target;

        var idx = target.lastIndexOf ('\\');
        var dir =  target.substr (1, idx);
        var name = target.substr (idx + 1, target.length);
        dir = remoteDrive + dir;
        WScript.Stdout.Write('File ' + name + ' start ' + start + ' finish ' + finish + ' duration ' + duration + ' modified ' + modified + '\n');

        if (!fso.FileExists(source)) {
            var msg = 'Source file does not exist: skipping ' + source;
            WScript.Stdout.Write(msg);
            fh.WriteLine (msg);
        }
        else if (!fso.FileExists(target)) {
            if (start != 0 || finish != duration) {
                // Convert to file
                WScript.Stdout.Write('Would convert file ' + i + " Name "  + target + '\n');
                var newfile = ConvertThisTrack(source);
                WScript.Stdout.Write('NEWFILE ' + newfile + '\n');
                if (!fso.FolderExists(dir)) {
                    CreateDirs (dir);   
                }
                fso.CopyFile (newfile, target, true);
                fso.DeleteFile (newfile);
            }
            else {
                // only create the directory if it doesn't exist then build from the ground up. 
                if (!fso.FolderExists(dir)) {
                    CreateDirs (dir);   
                }

                // if file exists but the size is different then copy again since its bunk
                if (fso.FileExists(target)) {
                    var f = fso.GetFile(target);
                    if (f.Size != size) {
                        fso.CopyFile (source, target, true);
                        fh.WriteLine ('Copying file since its the wrong size ' + f.Size + ' remote size ' + size);
                    }
                }
                else {
                    // file doesn't exist copy it over
                    fso.CopyFile (source, target, true);
                    fh.WriteLine ('Copying file since its new ' + target);
                    WScript.Stdout.Write("Copying file " + i + ' Copying file to SAN: ' + target + '\n');
                }
                WScript.Stdout.Write("Copying file " + i + ' Copying file to SAN: ' + target + '\n');
            }

            // Set the modified time to match the source file: all copied files
            var objFolder = so.NameSpace(dir);
            var objFolderItem = objFolder.ParseName(name);
            objFolderItem.ModifyDate = modified;
        }
        fh.WriteLine ('Count ' + i + ' Artist: ' + artist + ' Album: ' + album);
    }
    // Copy the XML file over as the last thing
    var sxml = sourceDrive + ':' + xmlPath  
    var txml = remoteDrive + ':' + xmlPath  
    WScript.Stdout.Write("copying xml file to targer: " + txml);
    fso.CopyFile (sxml, txml, true);
    fh.Close(); 
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

function ConvertThisTrack(track) 
{
    var opStatus = iTunesApp.ConvertFile2(track);
    var name = opStatus.trackName;
    WScript.Stdout.Write('Converting file ' + name);
    while (opStatus.InProgress) 
    {
        WScript.Sleep (500);
        WScript.Stdout.Write('Sleeping ' + name);
    }
    var outfile = fso.GetFile(opStatus.Tracks.Item(1).Location);
    WScript.Stdout.Write('OUTFILE ' + outfile);
    var track = opStatus.Tracks.Item(1);
    track.Delete();
    WScript.Stdout.Write('Deleted ' + outfile);
    return outfile;
}
