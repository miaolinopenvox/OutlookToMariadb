-- --------------------------------------------------------
-- 主机:                           172.16.88.8
-- 服务器版本:                        5.5.5-10.7.3-MariaDB - mariadb.org binary distribution
-- 服务器操作系统:                      Win64
-- HeidiSQL 版本:                  8.0.0.4396
-- --------------------------------------------------------

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET NAMES utf8 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;

-- 导出 outlook 的数据库结构
CREATE DATABASE IF NOT EXISTS `outlook` /*!40100 DEFAULT CHARACTER SET utf8mb3 */;
USE `outlook`;

-- 导出  表 outlook.email 结构
CREATE TABLE IF NOT EXISTS `email` (
  `id` int(10) NOT NULL AUTO_INCREMENT,
  `folder` varchar(200) DEFAULT NULL,
  `storeid` varchar(255) DEFAULT NULL,
  `bcc` varchar(500) DEFAULT NULL,
  `attachments` text DEFAULT NULL,
  `body` mediumtext DEFAULT NULL,
  `bodyformat` varchar(50) DEFAULT NULL,
  `cc` varchar(500) DEFAULT NULL,
  `creationtime` datetime DEFAULT NULL,
  `deferreddeliverytime` datetime DEFAULT NULL,
  `entryid` varchar(255) DEFAULT NULL,
  `htmlbody` mediumtext DEFAULT NULL,
  `importance` varchar(20) DEFAULT NULL,
  `internetcodepage` bigint(20) DEFAULT NULL,
  `lastmodificationtime` datetime DEFAULT NULL,
  `messageclass` varchar(255) DEFAULT NULL,
  `readreceiptrequested` bit(1) DEFAULT NULL,
  `receivedbyentryid` varchar(255) DEFAULT NULL,
  `receivedbyname` varchar(255) DEFAULT NULL,
  `receivedonbehalfofentryid` varchar(255) DEFAULT NULL,
  `receivedtime` datetime DEFAULT NULL,
  `recipients` mediumtext DEFAULT NULL,
  `replyrecipients` mediumtext DEFAULT NULL,
  `rtfbody` mediumtext DEFAULT NULL,
  `senderemailaddress` varchar(255) DEFAULT NULL,
  `senderemailtype` varchar(255) DEFAULT NULL,
  `sendername` varchar(255) DEFAULT NULL,
  `sentonbehalfOfName` mediumtext DEFAULT NULL,
  `size` bigint(20) DEFAULT NULL,
  `subject` varchar(500) DEFAULT NULL,
  `to` mediumtext DEFAULT NULL,
  `unread` bit(1) DEFAULT NULL,
  `action` varchar(1024) DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `storeid_entryid` (`storeid`,`entryid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;

-- 数据导出被取消选择。


-- 导出  表 outlook.folders 结构
CREATE TABLE IF NOT EXISTS `folders` (
  `name` varchar(255) NOT NULL,
  `storeid` varchar(255) NOT NULL,
  `entryid` varchar(255) NOT NULL,
  PRIMARY KEY (`name`),
  UNIQUE KEY `name` (`name`),
  UNIQUE KEY `storeid_entryid` (`storeid`,`entryid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb3;

-- 数据导出被取消选择。

-- 导出  视图 outlook.fulltext_email 结构
DROP VIEW IF EXISTS `fulltext_email`;
-- 创建临时表以解决视图依赖性错误
CREATE TABLE `fulltext_email` (
	`id` INT(10) NOT NULL,
	`storeid` VARCHAR(255) NULL COLLATE 'utf8mb3_general_ci',
	`entryid` VARCHAR(255) NULL COLLATE 'utf8mb3_general_ci',
	`fulltexts` LONGTEXT NULL COLLATE 'utf8mb3_general_ci'
) ENGINE=MyISAM;


-- 导出  视图 outlook.fulltext_email 结构
DROP VIEW IF EXISTS `fulltext_email`;
-- 移除临时表并创建最终视图结构
DROP TABLE IF EXISTS `fulltext_email`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`%` VIEW `outlook`.`fulltext_email` AS SELECT email.id, email.storeid, email.entryid, 
concat_ws(',',email.bcc, email.attachments,email.body,email.cc,email.entryid,
email.receivedbyentryid,email.receivedbyname,email.receivedonbehalfofentryid,
email.recipients,email.replyrecipients,email.senderemailaddress, email.senderemailtype,
email.sendername,email.sentonbehalfOfName,email.subject,email.`to`) as `fulltexts` from email ;
/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
