
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

Simply clone this repository to some drive, and also clone the Common repository to the same drive.

Note that currently all of my repositories expect to be cloned into the root of a drive so that they are 1) siblings, and 2) interrelated artifacts are found at: \Common....

In the future, I might implement a simple environment variable so that you can clone these off of the root (though they would still have to be siblings), but for now, they need to be at the root of a drive.

(So - you could SUBST a drive if you want to put them elsewhere)

