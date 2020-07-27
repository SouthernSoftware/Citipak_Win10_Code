
DEFINT A-Z
DECLARE SUB FixAPTrans ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, R%, c%, clr%)
DECLARE FUNCTION QPTrim$ (x$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB ReLinkTrans ()
DECLARE SUB RelinkBgtTrans ()
DECLARE SUB ChangeBgtDate ()



CONST False = 0, True = NOT False

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'AP.bi'

  
'changedate
'setfieldvalue

  LOCATE 1, 1, 0
  TextCursor -2, -2             'Define the Mouse cursor
  ShowCursor    'Show it
  
        FixAPTrans

END

SUB FixAPTrans
  CLS
   REDIM APTrans(1) AS APLedger81RecType
  RecLen = LEN(APTrans(1))
  APledgerFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APledgerFile LEN = RecLen
  NumTrans = LOF(APledgerFile) \ RecLen
  D$ = "10-09-97"

   Record! = 3242
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28038
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)

   Record! = 2630
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28039
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)

   Record! = 2987
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28040
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)

   Record! = 2977
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28041
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2857
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28042
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2860
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28043
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2864
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28044
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2976
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28045
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2978
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28046
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2863
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28047
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2865
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28048
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2981
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28048
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2973
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28049
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2866
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28050
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2979
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28050
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2982
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28050
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2946
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28051
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2557
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28052
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2625
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28053
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2969
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28053
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2900
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28054
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2868
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28055
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2980
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28056
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2628
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28057
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2966
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28058
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2942
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28059
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2634
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28060
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2870
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28061
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2869
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28062
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2878
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28063
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2939
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28064
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2964
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28065
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2965
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28066
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2871
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28067
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2646
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28068
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2872
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28069
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2873
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28070
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2643
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28071
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2875
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28071
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2990
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28072
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2967
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28073
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2874
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28074
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2876
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28075
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2970
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28076
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2945
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28077
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2963
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28078
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2644
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28079
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2570
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28080
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2571
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28080
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2858
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28080
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2940
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28081
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2941
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28082
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2944
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28083
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2968
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28083
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2961
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28084
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2960
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28085
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2959
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28086
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2879
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28087
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2880
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28088
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2881
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28088
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2989
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28089
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2972
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28090
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2957
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28091
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2958
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28092
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2986
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28093
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2882
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2883
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2884
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2885
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2886
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2887
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2888
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2889
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2890
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2891
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2893
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2894
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2896
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28094
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2975
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28095
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2901
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28096
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2902
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28097
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2653
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28098
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2657
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28098
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 26904
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28098
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2907
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28098
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2905
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28098
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2943
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28099
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2954
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28100
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2955
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28100
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2956
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28101
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2903
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28102
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2910
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28103
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2911
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28103
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2952
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28104
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2912
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28105
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2898
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28106
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2971
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28107
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2953
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28108
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2985
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28109
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2607
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28110
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2914
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28110
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2947
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28110
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2663
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28111
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2950
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28112
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2974
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28112
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2951
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28113
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2909
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28114
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2913
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28115
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2924
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28116
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2925
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28116
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2949
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28116
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2962
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28117
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2917
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28118
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2916
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28118
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2915
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28118
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2923
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28119
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2948
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28119
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2983
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28120
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2984
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28121
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   Record! = 2988
   GET APledgerFile, Record!, APTrans(1)
   APTrans(1).PDCheckNum = 28122
   APTrans(1).PDCheckDate = Date2Num%(D$)
   APTrans(1).PAYCODE = 3
   PUT APledgerFile, Record!, APTrans(1)
   CLOSE
  

END SUB

