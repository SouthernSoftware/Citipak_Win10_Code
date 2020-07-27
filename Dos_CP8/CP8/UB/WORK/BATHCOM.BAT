@echo off
rem ****************************************************
rem       Batch file to interface to CMDHCOM.EXE
rem       to execute type the following :-
rem 
rem             BatHcom [command(s)]
rem 
rem 
rem       Copyright Husky Computers Limited 1992.
rem                  July 1992.
rem
rem  NOTE: Use ':' instead of '=' in command line since
rem        the batch processing of the command line  by
rem        DOS thinks that '=' are command line separators.
rem
rem       Use        :- BATHCOM /TX:*.EXE
rem       Instead of :- BATHCOM /TX=*.EXE
rem
rem ****************************************************

:start

cmdhcom %1 %2 %3 %4 %5 %6 %7 %8 %9 /NOSCREEN

if ERRORLEVEL   1 echo ÿ
if ERRORLEVEL   1 echo A problem has occurred...

if ERRORLEVEL  99 goto 99bad_dos_version             
if ERRORLEVEL  28 goto 28cannot_find_file            
if ERRORLEVEL  27 goto 27cannot_attach_to_directory  
if ERRORLEVEL  26 goto 26cannot_log_onto_drive       
if ERRORLEVEL  25 goto 25unable_to_delete_file       
if ERRORLEVEL  24 goto 24unable_to_write_file        
if ERRORLEVEL  23 goto 23insufficient_remote_memory  
if ERRORLEVEL  22 goto 22user_hit_escape             
if ERRORLEVEL  21 goto 21hht_unable_to_create_file   
if ERRORLEVEL  20 goto 20could_not_find_handheld     
if ERRORLEVEL  19 goto 19free_space_error            
if ERRORLEVEL  18 goto 18dir_error                   
if ERRORLEVEL  17 goto 17list_drives_error           
if ERRORLEVEL  16 goto 16could_not_get_path          
if ERRORLEVEL  15 goto 15invalid_irq_or_addr         
if ERRORLEVEL  14 goto 14comms_after_download        
if ERRORLEVEL  13 goto 13general_failure             
if ERRORLEVEL  12 goto 12comms_failure               
if ERRORLEVEL  11 goto 11unable_to_read_file         
if ERRORLEVEL  10 goto 10unable_to_open_file         
if ERRORLEVEL   9 goto 9insufficient_memory         
if ERRORLEVEL   8 goto 8binary_file_not_found       
if ERRORLEVEL   7 goto 7invalid_handheld_type       
if ERRORLEVEL   6 goto 6invalid_filename            
if ERRORLEVEL   5 goto 5invalid_baud_rate           
if ERRORLEVEL   4 goto 4invalid_irq_line            
if ERRORLEVEL   3 goto 3invalid_port_number         
if ERRORLEVEL   2 goto 2invalid_config_path         
if ERRORLEVEL   1 goto 1invalid_parameter           
if ERRORLEVEL   0 goto 0ok                          



:99bad_dos_version             rem   99
     echo Sorry, this program requires DOS 3.x or greater.
     goto exit

:28cannot_find_file            rem   28
    echo Cannot find file.
    goto exit

:27cannot_attach_to_directory  rem   27
    echo Cannot attach to directory.
    goto exit

:26cannot_log_onto_drive       rem   26
    echo Cannot log onto drive.
    goto exit

:25unable_to_delete_file       rem   25
    echo Unable to delete file.
    goto exit

:24unable_to_write_file        rem   24
    echo Unable to write to file.
    goto exit

:23insufficient_remote_memory  rem   23
    Not enough memory on remote.
    goto exit

:22user_hit_escape             rem   22
    echo User aborted !!!!
    goto exit

:21hht_unable_to_create_file   rem   21
    echo Remote unable to create file.
    goto exit

:20could_not_find_handheld     rem   20
    echo No remote attached - please check connection.
    goto exit

:19free_space_error            rem   19
    echo Cannot calculate free space on handheld.
    goto exit

:18dir_error                   rem   18
    echo Invalid directory.
    goto exit

:17list_drives_error           rem   17
    echo Cannot list all drives.
    goto exit

:16could_not_get_path          rem   16
    echo Invalid path received.
    goto exit

:15invalid_irq_or_addr         rem   15
    echo Must have valid IRQ line and base address
    goto exit

:14comms_after_download        rem   14
    echo Cannot do communications after downloading software.
    goto exit

:13general_failure             rem   13
    echo General failure.
    goto exit

:12comms_failure               rem   12
    echo Communications error - please try again.
    goto exit

:11unable_to_read_file         rem   11
    echo Unable to read file.
    goto exit

:10unable_to_open_file         rem   10
    echo Unable to open file.
    goto exit

:9insufficient_memory         rem    9
    echo Insufficient memory
    goto exit

:8binary_file_not_found       rem    8
    echo Cannot find software download file for handheld.
    goto exit

:7invalid_handheld_type       rem    7
    echo Invalid handheld type, must be FS2, HUNTER16, HAWK or HUNTER
    goto exit

:6invalid_filename            rem    6
    echo Invalid filename.
    goto exit

:5invalid_baud_rate           rem    5
    echo Invalid baud rate entered, must be one of :-
    echo        38400 for FS/2, Hunter 16 or Hunter 16/80.
    echo        19200 for Hawk.
    echo         4800 for Hunter or Hunter 2.
    goto exit

:4invalid_irq_line            rem    4
    echo Invalid IRQ line.
    goto exit

:3invalid_port_number         rem    3
    echo Invalid port number entered.
    goto exit

:2invalid_config_path         rem    2
    echo Cannot find configuration file on that path.
    goto exit

:1invalid_parameter           rem    1
    echo Invalid parameter entered.
    echo Type CMDHCOM /? for options.
    goto exit

:0ok                          rem    0

:exit
echo ÿ
