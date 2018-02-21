
# The InnoVisioNate Properties Component
----
This component manages the storage (persistence) and property pages implementation for any property in your system.

By "ANY" we mean anything at all, variables, objects, even the properties of other COM objects.

This is a product I originally developed in the 1998 timeframe. It was essentially complete back then, though a few years ago I implemented the more "modern" style of property sheets.

It has been rock solid ever since. I use it in pretty much everything I do to remember everything about the state of a system to restore it at next runtime.

In fact, creating property sheets is easy enough, that you can use them rather than dialogs you would have to write to manage your variables and data.

Full documentation on using the system is provided in the Documentation directory.

## Building the Properties component

----

Simply clone this repository somwhere on a drive, and also clone the Common repository as a sibling of this repository on the same drive

I have added the "GSYSTEM_HOME" environment variable to allow you to specify where on the system these repositories are - note that this definition is in the "Compiler Options.props" file in the Common Repository.

Detailed instructions for using the environment variable are given in the Common Repository.

