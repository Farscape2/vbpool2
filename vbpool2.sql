-- phpMyAdmin SQL Dump
-- version 4.6.6deb5
-- https://www.phpmyadmin.net/
--
-- Host: localhost:3306
-- Gegenereerd op: 28 apr 2020 om 11:17
-- Serverversie: 10.3.22-MariaDB-0+deb10u1
-- PHP-versie: 7.3.14-1~deb10u1

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: vbpool2
--

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblAddress
--

CREATE TABLE tblAddress (
  addressID int(11) NOT NULL DEFAULT 0,
  firstname varchar(50) DEFAULT NULL,
  middlename varchar(32) DEFAULT NULL,
  lastname varchar(50) DEFAULT NULL,
  shortname varchar(24) DEFAULT NULL,
  address varchar(50) DEFAULT NULL,
  postalcode varchar(10) DEFAULT NULL,
  city varchar(50) DEFAULT NULL,
  telephone varchar(20) DEFAULT NULL,
  email varchar(255) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblCompetitors
--

CREATE TABLE tblCompetitors (
  competitorID int(11) NOT NULL,
  poolid int(11) NOT NULL,
  addressID int(11) DEFAULT NULL,
  nickName varchar(50) NOT NULL,
  payed tinyint(1) DEFAULT 0,
  predictionTeam1 int(11) DEFAULT NULL,
  predictionTeam2 int(11) DEFAULT NULL,
  predictionTeam3 int(11) DEFAULT NULL,
  predictionTeam4 int(11) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblPoolPoints
--

CREATE TABLE tblPoolPoints (
  poolid int(11) NOT NULL,
  pointTypeID int(11) NOT NULL,
  pointPointsAward int(11) DEFAULT 0,
  pointPointsMargin tinyint(3) UNSIGNED DEFAULT 0
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblPools
--

CREATE TABLE tblPools (
  poolID int(11) NOT NULL DEFAULT 0,
  tournamentID int(11) DEFAULT NULL,
  organisationID int(11) DEFAULT NULL,
  poolName varchar(50) DEFAULT NULL,
  poolStartAcceptForms datetime DEFAULT NULL,
  poolEndAcceptForms datetime DEFAULT NULL,
  poolCost decimal(19,4) DEFAULT 10.0000,
  prizeHighDayScore decimal(19,4) DEFAULT 0.0000,
  prizeHighDayOverallPosition decimal(19,4) DEFAULT 0.0000,
  prizeLowDayOverallPosition decimal(19,4) DEFAULT 0.0000,
  prizePercentageFirst double DEFAULT 0,
  prizePercentageSecond double DEFAULT 0,
  prizePercentageThird double DEFAULT 0,
  prizePercentageFourth double DEFAULT 0,
  prizeLowFinalOverallPosition decimal(19,4) DEFAULT 0.0000
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblPredictionGroupResults
--

CREATE TABLE tblPredictionGroupResults (
  competitorID int(11) NOT NULL,
  groupLetter varchar(1) DEFAULT NULL,
  predictionGroupPosition1 varchar(255) DEFAULT NULL,
  predictionGroupPosition2 varchar(255) DEFAULT NULL,
  predictionGroupPosition3 varchar(255) DEFAULT NULL,
  predictionGroupPosition4 varchar(255) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblPredictionTopscorers
--

CREATE TABLE tblPredictionTopscorers (
  competitorID int(11) NOT NULL,
  predictionTopscorerPosittion int(11) DEFAULT NULL,
  predictionTopscorePlayerID int(11) DEFAULT NULL,
  predictionTopscoreGoals int(11) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblPrediction_Finals
--

CREATE TABLE tblPrediction_Finals (
  competitorID int(11) NOT NULL,
  matchNumber int(11) DEFAULT NULL,
  teamNameA varchar(255) DEFAULT NULL,
  teamNameB varchar(255) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblPrediction_MatchResults
--

CREATE TABLE tblPrediction_MatchResults (
  competitorID int(11) DEFAULT NULL,
  matchNumber int(11) DEFAULT NULL,
  predictionGoalsHalftimeA tinyint(3) UNSIGNED DEFAULT NULL,
  predictionGoalsHalftimeB tinyint(3) UNSIGNED DEFAULT 0,
  predictionGoalsFulltimeA tinyint(3) UNSIGNED DEFAULT 0,
  predictionGoalsFulltimeB tinyint(3) UNSIGNED DEFAULT 0,
  predictionResultToto tinyint(3) UNSIGNED DEFAULT NULL,
  pnt int(11) DEFAULT 0
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblPrediction_Numbers
--

CREATE TABLE tblPrediction_Numbers (
  competitorID int(11) NOT NULL,
  predictionTypeID int(11) DEFAULT NULL,
  predictionNumber int(11) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- --------------------------------------------------------

--
-- Tabelstructuur voor tabel tblUsers
--

CREATE TABLE tblUsers (
  ID int(11) NOT NULL DEFAULT 0,
  UserName varchar(25) NOT NULL,
  Password varchar(32) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

--
-- Indexen voor geëxporteerde tabellen
--

--
-- Indexen voor tabel tblAddress
--
ALTER TABLE tblAddress
  ADD PRIMARY KEY (addressID);

--
-- Indexen voor tabel tblCompetitors
--
ALTER TABLE tblCompetitors
  ADD KEY competitorID (competitorID);

--
-- Indexen voor tabel tblPools
--
ALTER TABLE tblPools
  ADD PRIMARY KEY (poolID);

--
-- AUTO_INCREMENT voor geëxporteerde tabellen
--

--
-- AUTO_INCREMENT voor een tabel tblCompetitors
--
ALTER TABLE tblCompetitors
  MODIFY competitorID int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=518;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
