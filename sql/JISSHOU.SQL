
 
 CREATE TABLE BL (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        BLNo                 char(20) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        ShipLine             char(5) NOT NULL,
        RFFlag               char(1) NOT NULL,
        DGFlag               char(1) NOT NULL,
        LPort                char(5) NOT NULL,
        DPort                char(5) NOT NULL,
        TSPort               char(5) NOT NULL,
        DeliverTime          datetime NULL,
        RecTerminal          char(20) NOT NULL,
        PRIMARY KEY (VslCode, VoyCtrl, BLNo)
 )
go
 
 
 CREATE TABLE Booking (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        BookNo               char(20) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        TransCode            char(4) NOT NULL,
        ComTime              datetime NULL,
        Sender               char(20) NOT NULL,
        TransPlace           char(20) NOT NULL,
        ShipLine             char(5) NOT NULL,
        RFFlag               char(1) NOT NULL,
        DGFlag               char(1) NOT NULL,
        LPort                char(5) NOT NULL,
        DPort                char(5) NOT NULL,
        TSPort               char(5) NOT NULL,
        Qty20                int NULL,
        Qty40                int NULL,
        QtyOth               int NULL,
        RecTerminal          char(20) NOT NULL,
        PRIMARY KEY (VslCode, VoyCtrl, BookNo)
 )
go
 
 
 CREATE TABLE Container (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        ContNo               char(12) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        ContSize             char(2) NOT NULL,
        ContType             char(2) NOT NULL,
        ContHeight           char(2) NOT NULL,
        ContWeight           int NULL,
        SealNo               char(15) NOT NULL,
        ShipLine             char(5) NOT NULL,
        ListDate             datetime NULL,
        ListNo               char(15) NOT NULL,
        LatestTransCode      char(4) NOT NULL,
        OffDockFlag          char(1) NOT NULL,
        StackPlace           char(20) NOT NULL,
        PRIMARY KEY (VslCode, VoyCtrl, ContNo)
 )
go
 
 
 CREATE TABLE ContTransaction (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        ContNo               char(12) NOT NULL,
        TransCode            char(4) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        ComTime              datetime NULL,
        Sender               char(20) NOT NULL,
        TransPlace           char(20) NOT NULL,
        TransData            char(100) NOT NULL,
        PRIMARY KEY (VslCode, VoyCtrl, ContNo, TransCode)
 )
go
 
 
 CREATE TABLE ExportCont (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        ContNo               char(12) NOT NULL,
        BookNo               char(20) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        RFFlag               char(1) NOT NULL,
        DGFlag               char(1) NOT NULL,
        PickPlace            char(20) NOT NULL,
        CustOK               char(1) NOT NULL,
        ExpPermitDate        datetime NULL,
        ExpPermitNo          char(15) NOT NULL,
        BLNo                 char(20) NOT NULL,
        EmpDelTime           datetime NULL,
        WHArTime             datetime NULL,
        VanTime              datetime NULL,
        WHDelTime            datetime NULL,
        ExpArTime            datetime NULL,
        OutTime              datetime NULL,
        PRIMARY KEY (VslCode, VoyCtrl, ContNo, BookNo)
 )
go
 
 
 CREATE TABLE ImportCont (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        ContNo               char(12) NOT NULL,
        BLNo                 char(20) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        RFFlag               char(1) NOT NULL,
        DGFlag               char(1) NOT NULL,
        HotFlag              char(1) NOT NULL,
        EstDelDate           datetime NULL,
        FreeTime             datetime NULL,
        OLTICFlag            char(1) NOT NULL,
        OLTICDate            datetime NULL,
        OLTICNo              char(15) NOT NULL,
        DOStatus             char(1) NOT NULL,
        DelPermitDate        datetime NULL,
        OLTDateFrom          datetime NULL,
        OLTDateTo            datetime NULL,
        ReturnPlace          char(20) NOT NULL,
        InTime               datetime NULL,
        CYDelTime            datetime NULL,
        SYDelTime            datetime NULL,
        WHArDate             datetime NULL,
        DevanTime            datetime NULL,
        ReturnTime           datetime NULL,
        PRIMARY KEY (VslCode, VoyCtrl, ContNo, BLNo)
 )
go
 
 
 CREATE TABLE InspectCont (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        ContNo               char(12) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        BotStatus            char(1) NOT NULL,
        ZooStatus            char(1) NOT NULL,
        CustStatus           char(1) NOT NULL,
        PRIMARY KEY (VslCode, VoyCtrl, ContNo)
 )
go
 
 
 CREATE TABLE mPort (
        PortCode             char(5) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             char(18) NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        FullName             char(20) NOT NULL,
        PRIMARY KEY (PortCode)
 )
go
 
 
 CREATE TABLE mShipLine (
        ShipLine             char(5) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        FullName             char(20) NOT NULL,
        PRIMARY KEY (ShipLine)
 )
go
 
 
 CREATE TABLE mVessel (
        VslCode              char(7) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        FullName             char(20) NOT NULL,
        PRIMARY KEY (VslCode)
 )
go
 
 
 CREATE TABLE Pickup (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        BookNo               char(20) NOT NULL,
        PickPlace            char(20) NOT NULL,
        ContSize             char(2) NOT NULL,
        ContType             char(2) NOT NULL,
        ContHeight           char(2) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        Qty                  int NULL,
        TransCode            char(4) NOT NULL,
        ComTime              datetime NULL,
        Sender               char(20) NOT NULL,
        TransPlace           char(20) NOT NULL,
        PRIMARY KEY (VslCode, VoyCtrl, BookNo, PickPlace, ContSize, ContType, ContHeight)
 )
go
 
 
 CREATE TABLE sTransaction (
        ComTime              datetime NOT NULL,
        SeqNo                int NOT NULL,
        TransCode            char(4) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NULL,
        Process              char(1) NOT NULL,
        Sender               char(20) NOT NULL,
        TransPlace           char(20) NOT NULL,
        Item01               char(20) NOT NULL,
        Item02               char(20) NOT NULL,
        Item03               char(20) NOT NULL,
        Item04               char(20) NULL,
        Item05               char(20) NULL,
        Item06               char(20) NULL,
        Item07               char(20) NULL,
        Item08               char(20) NULL,
        Item09               char(20) NULL,
        Item10               char(20) NULL,
        Item11               char(20) NULL,
        Item12               char(20) NULL,
        Item13               char(20) NULL,
        Item14               char(20) NULL,
        Item15               char(20) NULL,
        Item16               char(20) NULL,
        Item17               char(20) NULL,
        Item18               char(20) NULL,
        Item19               char(20) NULL,
        Item20               char(20) NULL,
        PRIMARY KEY (ComTime, SeqNo, TransCode)
 )
go
 
 
 CREATE TABLE Terminal (
        Terminal             char(2) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NULL,
        UpdtTmnl             char(15) NOT NULL,
        YardWaitTime         int NULL,
        GateWaitTime         int NULL,
        ComTime              datetime NULL,
        Sender               char(20) NOT NULL,
        TransPlace           char(20) NULL,
        PRIMARY KEY (Terminal)
 )
go
 
 
 CREATE TABLE VslPort (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        PortCode             char(5) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        CallSeq              int NULL,
        ETA                  datetime NULL,
        TA                   datetime NULL,
        ETD                  datetime NULL,
        TD                   datetime NULL,
        PRIMARY KEY (VslCode, VoyCtrl, PortCode)
 )
go
 
 
 CREATE TABLE VslSchedule (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        ShipLine             char(5) NOT NULL,
        DsVoyage             char(12) NOT NULL,
        LdVoyage             char(12) NOT NULL,
        CYOpen               datetime NULL,
        CYCut                datetime NULL,
        PRIMARY KEY (VslCode, VoyCtrl)
 )
go
 
 
 CREATE TABLE VslTransaction (
        VslCode              char(7) NOT NULL,
        VoyCtrl              int NOT NULL,
        TransCode            char(4) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        ComTime              datetime NULL,
        Sender               char(20) NOT NULL,
        TransPlace           char(20) NOT NULL,
        TransData            char(100) NOT NULL,
        PRIMARY KEY (VslCode, VoyCtrl, TransCode)
 )
go
 
 