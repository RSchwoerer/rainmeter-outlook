# Basic Structure #

Create a measure that references the plugin
```
[MeasureOutlook]
Measure=Plugin
Plugin=Plugins\OutlookPlugin.dll
```

In a measure, you can specify the following options:
  * **Resource** (required) Define the resource you want to access. This has to be one of the resources listed below, or another measures [(example)](ReferencesInResource.md).
  * **Result** Specify the value you want in the form _%PropertyName_. If the result should be a string, you can write something more complex, like _%UnreadItemCount new mails in %Name_. The available properties depend on the resource type and are listed below.
  * **Default** If you access a resource that does not exist (i.e. when the index is out of bounds), this value will be returned instead.
  * **OnError** If something goes wrong, this value will be used. You can use _%Message_ to get details.

Additionally, you can use these options to modify the result:
  * **Select** selects a logical sub-group.
  * **Filter** keeps all items for which the value evaluates to 1. The value has to be a _%PropertyName_.

# Resources #
This section lists available resources and their properties.

## Status ##
Use this to check if the plugin is connected.

**Options**
  * **OkMessage** overrides the status message if the status is OK.

**Properties**
  * **%Code** A value < 0 on errors, and > 0 if the plugin is connected.
  * **%IsOk** when the plugin is connect 1, otherwise 0.
  * **%Message** The status message.

## MAPIFolder ##
Collects all folders in your inbox. The result is a folder list.

**Options**
  * **Root** (only if **Resource** does not reference another measure) Select the roots from which the folders are collected. See [Roots](Roots.md) for more information.

### Folder list ###
A list of MAPI folders.

**Select**
  * **Root** selects only the root folders. (There are none if the result has been filtered before).

**Options**
  * **Index** selects a single folder of the list. In this case, the resource type is not a folder list, but a single folder.

**Properties**
  * **%Count** number of folders
  * **%TotalUnreadItemCount** total number of unread items

### Folder ###
A single MAPI folder.

**Select**
  * **Subfolders** selects all (transitive) subfolders. In this case, the resource type changes to folder list.

**Properties**
  * **%ItemCount** number of items
  * **%HasUnreadItems** when the folder contains unread items 1, otherwise 0
  * **%Name**
  * **%Path**
  * **%TotalUnreadItemCount** number of unread items in this folder and its subfolders
  * **%UnreadItemCount** number of unread items in this folder