
use if0_35817478_foo; 

/*
DROP TABLE IF EXISTS `userz`; 
DROP TABLE IF EXISTS `userz0`; 
*/

/*
CREATE TABLE `userz0` (
   `userzID` int NOT NULL AUTO_INCREMENT,
   `uName` varchar(64) NOT NULL DEFAULT 'ELMERFUDD',
   `pw` varchar(128) NOT NULL DEFAULT '9876543210',
   `pwHash` varchar(128) NOT NULL DEFAULT '9876543210',
   `fName` varchar(64) NOT NULL DEFAULT 'ELMER',
   `lName` varchar(64) NOT NULL DEFAULT 'FUDD',
   `active` int NOT NULL DEFAULT '-1',
   `test` int NOT NULL DEFAULT '-1',
   `descr` varchar(2048) NOT NULL DEFAULT '',
--   / *
   `varsServer` varchar(2048) NOT NULL DEFAULT '',
   `varsClient` varchar(2048) NOT NULL DEFAULT '',
   `email0` varchar(256) NOT NULL DEFAULT '',
   `email1` varchar(256) NOT NULL DEFAULT '',
   `email2` varchar(256) NOT NULL DEFAULT '',
   `phone0` varchar(32) NOT NULL DEFAULT '',
   `phone1` varchar(32) NOT NULL DEFAULT '',
   `phone2` varchar(32) NOT NULL DEFAULT '',
   `address0` varchar(256) NOT NULL DEFAULT '',
   `address1` varchar(256) NOT NULL DEFAULT '',
   `address2` varchar(256) NOT NULL DEFAULT '',
   `someOtherNote` varchar(2048) NOT NULL DEFAULT '',
   `devNote` varchar(64) NOT NULL DEFAULT 'Sandbox and should be broken out to multiple related tables.',
--   * /
   PRIMARY KEY (`userzID`)
 ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
 */
 
 #Error Code: 1118. 
 
 
  /*
  
  Row size too large. The maximum row size for the used table type, 
  not counting BLOBs, is 65535. This includes storage overhead, check the manual. 
  You have to change some columns to TEXT or BLOBs 0.156 sec
  
  	Error Code: 1064. You have an error in your SQL syntax; check the manual 
  that corresponds to your MySQL server version for the right syntax 
  to use near 'Error Code: 1118. 
  */
 
/* 
INSERT INTO `if0_35817478_foo`.`userz0`
(`userzID`)
VALUES
(999);
*/ 
INSERT INTO `if0_35817478_foo`.`userz0`
(
`uName`,
`pw`,
`pwHash`,
`fName`,
`lName`,
`email0`
)
SELECT  
 'JKSFO_uName'
, 'JKSFO_pw'
, 'JKSFO_pwHash'
, 'JKSFO_fName'
, 'JKSFO_lName'
, 'littermates828@gmail.com'
UNION 
SELECT  
 'JKSFO0_uName'
, 'JKSFO0_pw'
, 'JKSFO0_pwHash'
, 'JKSFO0_fName'
, 'JKSFO0_lName'
, 'littermates828@gmail.com'
 
UNION 
SELECT  
 'JKSFO1_uName'
, 'JKSFO1_pw'
, 'JKSFO1_pwHash'
, 'JKSFO1_fName'
, 'JKSFO1_lName'
, 'littermates828@gmail.com'
 ;

# select * from `if0_35817478_foo`.`userz0`
