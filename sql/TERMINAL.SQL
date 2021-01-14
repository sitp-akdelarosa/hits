 
 CREATE TABLE Terminal (
        Terminal             char(2) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NULL,
        UpdtTmnl             char(15) NOT NULL,
        RecWaitTime          int NULL,
        DelWaitTime          int NULL,
        RDWaitTime           int NULL,
        GateWaitTime         int NULL,
        ComTime              datetime NULL,
        Sender               char(20) NOT NULL,
        TransPlace           char(20) NULL,
        PRIMARY KEY (Terminal)
 )
go
