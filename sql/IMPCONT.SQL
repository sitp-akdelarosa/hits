
 
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
        CustStatus           char(1) NOT NULL,
        FreeTimeExt          datetime NULL,
        PRIMARY KEY (VslCode, VoyCtrl, ContNo, BLNo)
 )
go
 
 