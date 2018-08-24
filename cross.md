# VBA Project: vbaGit_Word
This cross reference list for repo (vbaGit_Word) was automatically created on 24/08/2018 18:03:54 by VBAGit.For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")
You can see [library and dependency information here](dependencies.md)

###Below is a cross reference showing which modules and procedures reference which others
*module*|*proc*|*referenced by module*|*proc*
---|---|---|---
cJobject||VbaGit|getVGSettings
cJobject||VbaGit|getAllInfoFiles
cJobject||VbaGit|getDependencyList
cJobject||VbaGit|dependencyResolve
cJobject||VbaGit|makeCrossReferenceJob
cJobject||VbaGit|getProjects
cJobject||VbaGit|makeInfoFile
cregXLib||regXLib|rxMakeRxLib
cStringChunker||VbaGit|dependencyResolve
cStringChunker||VbaGit|makeCross
cStringChunker||VbaGit|makeReadMe
cStringChunker||VbaGit|makeDependency
cStringChunker||VbaGit|constructModLink
cStringChunker||VbaGit|makeArguments
cStringChunker||VbaGit|findModLink
cStringChunker||VbaGit|getDependencyList
cVBAArgument||cVBAProcedure|dealWithArguments
cVbaGit||VbaGit|doImportFromGit
cVbaGit||VbaGit|doGit
cVBAmodule||VbaGit|isModuleObj
cVBAmodule||VbaGit|getmoduleList
cVBAProcedure||VbaGit|getProcList
cVBAProject||VbaGit|getProjects
regXLib|rxReplace|usefulcJobject|hackJSONPObjectToJSON
regXLib|rxReplace|usefulcJobject|cleanGoogleWire
regXLib|rxReplace|usefulcJobject|hackJSObjectToJSON
urlResult|urlGet|cVbaGit|getSpecificRepo
urlResult|urlGet|cVbaGit|getFileByPath
urlResult|urlGet|cVbaGit|getUnpaged
urlResult|urlGet|cVbaGit|getMyRepos
urlResult|urlPost|cVbaGit|commitFile
urlResult|urlPost|cVbaGit|getTokenFromBasic
urlResult|urlPost|cVbaGit|createRepo
usefulcJobject|JSONParse|VbaGit|doImportFromGit
usefulcJobject|JSONParse|VbaGit|getAllInfoFiles
usefulRegex|getRx|VbaGit|dependencyResolve
usefulRegex|getTheEndRx|VbaGit|testmodulestuff
usefulStuff|Base64Encode|VbaGit|setGitClientCredentials
usefulStuff|Base64Encode|VbaGit|setGitBasicCredentials
usefulStuff|checkOrCreateFolder|VbaGit|writeInfoFile
usefulStuff|conditionalAssignment|VbaGit|modulesToInfo
usefulStuff|getAllSubFolderPaths|VbaGit|getAllInfoFiles
usefulStuff|getTimestampFromDate|VbaGit|makeInfoFile
usefulStuff|isSomething|VbaGit|getDependencyList
usefulStuff|isSomething|VbaGit|dependencyResolve
usefulStuff|isSomething|VbaGit|makeCrossReferenceJob
usefulStuff|isSomething|VbaGit|createRepos
usefulStuff|isUndefined|VbaGit|getVGSettings
usefulStuff|readFromFolderFile|VbaGit|writeTheFiles
usefulStuff|writeToFolderFile|VbaGit|writeToStagingArea
