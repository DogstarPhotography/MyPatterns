Attribute VB_Name = "AnObjectModelDescription"
'-----------------------------------------------------------------------------
'   This is a text file containign a descritpion of the objects in ths class and how they are to be used
'-----------------------------------------------------------------------------
'   Engine Object methods and events
'
'   RunFile, RunText:
'       -
'   ImmediateRunFile, ImmediateRunText:
'       -Currently commented out
'   Scriptupdate, ScriptCompleted:
'       -
'
'-----------------------------------------------------------------------------
'   Script File Format Information
'
'   Script files must be in the following format(s):
'
'   [ACTION]|[IMPORT]|[EXPORT]|[PURGE]
'   server=SERVER_APPLICATION.SERVER_OBJECT
'   source=SOURCE_NAME
'   destination=DESTINATION_NAME
'   options=OPTIONS_ARE_SERVER_SPECIFIC_AND_OPTIONAL
'       note that there must be at least server and source data
'       and that these must follow the IServerScriptAbs interface
'
'   [SCHEDULE]
'   schedule=DATE_AND_TIME_TO_RUN
'       Note that the schedule action will cause the script to stop execution
'       at that point and it will be re-executed ENTIRELY at a later date/time
'       Thus the Schedule action is best used at the start of a script
'
'   [LAUNCH]
'   program=PROGRAM_COMMAND_LINE
'
'   [SLEEP]
'   sleep=TIME_TO_CONTINUE
'       The script will wait at this point until the designated time
'
'   [Log]
'   logfile=LOG_FILE_NAME
'
'   [NOLOG]
'
'   [FILE]
'   source=SOURCE_FILE_AND_OR_DIRECTORY
'   destination=OPTIONAL_DESTINATION_FILE_AND_OR_DIRECTORY
'   options=copy|move|delete|rename
'       only one of the above may be chosen
'
'   [DIRECTORY]
'   source=DIRECTORY
'   options=create|delete|rename
'       only one of the above may be chosen
'
'   [SCRIPT]
'   script=SCRIPT_FILE_TO_RUN
'       SCRIPT, SCHEDULE and SLEEP elements will be ignored in sub scripts
'       as they can cause too many errors
'
'   [END]
'
'-----------------------------------------------------------------------------

