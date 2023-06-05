# EA-Matic Trouble Shooting

## "An error occurred while attempting to communicate with an Addin"

When running an EA-Matic script it causes an error dialog with the message:

> An error occurred while attempting to communicate with an Addin:
>
> EA-Matic (EAScriptAddin.EAScriptAddingAddingClass_
>
> ...
>
> Please contact the add-in provided for assistance"

This is caused by EA-Matic not being able to run the VBScript file.

Potential causes of this error are:

* not installing the prequisites listed at [https://bellekens.com/ea-matic/](https://bellekens.com/ea-matic/)

* having runtime/compilation errors

* not being able to inline the included files

How to fix these issues:

EA-Matic requires [.Net framework 4.5](https://www.microsoft.com/en-us/download/details.aspx?id=17851).

Some older EA VB scripts will require [.Net framework 3.5](https://www.microsoft.com/en-au/download/details.aspx?id=25150)

Make your EA-Matic script a thin wrapper around a regular VBScript and make sure you can run that correctly in EA.

EA-Matic inlines all included files when it builds the VBScript to use at
runtime. All `!INC ...` statements are replaced with the contents of the
included file. However there is a known bug [EA-Matic: Fail with error when
getIncludedcode can not find key includeString in
includableScripts](https://github.com/GeertBellekens/Enterprise-Architect-Toolpack/issues/120).
The current way this is implemented when the included file can not be found is
to replace the `!INC` with an empty string. This will cause errors like `Variable is undefined`.

Until this gets fixed you have two work-around options:

1. Enable dev mode (Specialize > EA-Matics > settings > enable Developer Mode)

   This fixes the problem because EA-Matic will load all scripts into its cache.

   This is the preferred option. Only if you have thousands of scripts will this
   start to cause a performance issue.

1. Add `!EA-Matic` to all files are are included

   This fixes the problem because EA-Matic will load all scripts into its cache
   that include `EA-Matic` in the file. This work around is more tedious because
   you need to traverse the included file graph and modify every file found, and
   if you change any include files you need to keep this in sync.

## EA-Matic only saves work every 5 minutes

**This is by design.**

The thing is that EA-Matic will be triggered with each and every event in EA. If
we were to read and interpret all scripts on every event, that would slow down
EA considerably.

That's why EA-Matic keeps an in-memory copy of all EA-Matic functions. Every 5
minutes (or when you open the settings) EA-Matic will check if any changes have
been made to the EA-Matic scripts, and if so, it will refresh the scripts in
memory.

This can sometimes be annoying for the script developer, but it's the best
solution for the users.

## EA-Matic scripts are stale, old versions

When you run EA-Matic the scripts aren't doing what you just finished writing
and saved into you VBScript.

See [EA-Matic only saves work every 5 minutes](#ea-matic-only-saves-work-every-5-minutes)
