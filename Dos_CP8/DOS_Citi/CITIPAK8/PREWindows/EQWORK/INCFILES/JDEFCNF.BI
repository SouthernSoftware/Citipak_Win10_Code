TYPE Config
    Active        AS INTEGER     'Config initialized flag
    MonTyp        AS INTEGER     'Monitor Type = Monitor (NOT MonoCode)
    CsrSize       AS INTEGER     'Cursor Scan Lines
    Mouse         AS INTEGER     'Moose Flag
    PulBar        AS INTEGER     'Horizontal Bar Color
    MenBox        AS INTEGER     'Pulled down Box Color
    ActivCh       AS INTEGER     'Active menu choice color
    InActCh       AS INTEGER     'Inactive (Non-selectable) Choice Color
    HiLite        AS INTEGER     'Menu Bar Color for active choices
    InActHiLt     AS INTEGER     'Dito for inactive choices
    NonMen        AS INTEGER     'Normal screen color
    CPUSpeed      AS LONG        'Relative processor speed
END TYPE

TYPE jConfig
    MonTyp AS INTEGER           'Monitor Type
    PulBar AS INTEGER           'Horizontal Bar Color
    MenBox AS INTEGER           'Pulled down Box Color
    ActivCh AS INTEGER          'Active menu choice color
    InActCh AS INTEGER          'Inactive (Non-selectable) Choice Color
    HiLite AS INTEGER           'Menu Bar Color for active choices
    InActHiLt AS INTEGER        'Dito for inactive choices
    NonMen AS INTEGER           'Normal screen color
    CurSize AS INTEGER          'Cursor Scan Lines
    Mouse AS INTEGER            'Mouse in residence Flag
END TYPE

'COMMON SHARED Cnf AS Config

