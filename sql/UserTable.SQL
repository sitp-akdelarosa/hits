
 CREATE TABLE lUserTable (
        UserID               char(5) NOT NULL,
        timestamp            timestamp NULL,
        CompanyName          char(256) NOT NULL,
        MailAddress          char(256) NULL,
        BeforeCntnrNo        char(12) NULL,
        PRIMARY KEY (UserID)
 )
go
