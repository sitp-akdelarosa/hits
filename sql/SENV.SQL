
 CREATE TABLE sEnvironment (
        dummy                char(1) NOT NULL,
        timestamp            timestamp NULL,
        UpdtTime             datetime NULL,
        UpdtPgCd             char(8) NOT NULL,
        UpdtTmnl             char(15) NOT NULL,
        CurrentPort          char(5) NOT NULL,
        PRIMARY KEY (dummy)
 )
go
