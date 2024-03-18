-- phpMyAdmin SQL Dump
-- version 4.6.6deb5
-- https://www.phpmyadmin.net/
--
-- Servidor: localhost:3306
-- Tiempo de generación: 05-09-2020 a las 14:27:26
-- Versión del servidor: 10.3.22-MariaDB-0+deb10u1
-- Versión de PHP: 7.3.19-1~deb10u1

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `imperiumclassic`
--
CREATE DATABASE IF NOT EXISTS `imperiumclassic` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;
USE `imperiumclassic`;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `atributos`
--

CREATE TABLE `atributos` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `att1` tinyint(3) UNSIGNED NOT NULL,
  `att2` tinyint(3) UNSIGNED NOT NULL,
  `att3` tinyint(3) UNSIGNED NOT NULL,
  `att4` tinyint(3) UNSIGNED NOT NULL,
  `att5` tinyint(3) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `banco_items`
--

CREATE TABLE `banco_items` (
      `user_id` mediumint(8) UNSIGNED NOT NULL,
      `item_id1` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount1` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
      `item_id2` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount2` smallint(5) UNSIGNED NULL DEFAULT '0',
    
       `item_id3` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount3` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id4` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount4` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id5` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount5` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id6` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount6` smallint(5) UNSIGNED NULL DEFAULT '0',
   
       `item_id7` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount7` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id8` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount8` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id9` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount9` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id10` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount10` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id11` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount11` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id12` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount12` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id13` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount13` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id14` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount14` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id15` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount15` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id16` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount16` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id17` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount17` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id18` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount18` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id19` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount19` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id20` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount20` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id21` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount21` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id22` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount22` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id23` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount23` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id24` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount24` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id25` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount25` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id26` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount26` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id27` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount27` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id28` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount28` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id29` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount29` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id30` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount30` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id31` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount31` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id32` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount32` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id33` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount33` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id34` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount34` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id35` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount35` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id36` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount36` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id37` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount37` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id38` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount38` smallint(5) UNSIGNED NULL DEFAULT '0',
	  
	  `item_id39` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount39` smallint(5) UNSIGNED NULL DEFAULT '0',
    
      `item_id40` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount40` smallint(5) UNSIGNED NULL DEFAULT '0') ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `inventario_items`
--

CREATE TABLE inventario_items (
    `user_id` mediumint(8) UNSIGNED NOT NULL,
      `item_id1` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount1` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped1` tinyint(1) UNSIGNED NULL DEFAULT '0',
	  
      `item_id2` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount2` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped2` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
       `item_id3` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount3` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped3` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id4` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount4` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped4` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id5` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount5` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped5` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id6` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount6` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped6` tinyint(1) UNSIGNED NULL DEFAULT '0',
   
       `item_id7` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount7` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped7` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id8` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount8` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped8` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id9` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount9` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped9` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id10` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount10` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped10` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id11` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount11` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped11` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id12` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount12` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped12` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id13` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount13` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped13` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id14` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount14` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped14` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id15` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount15` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped15` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id16` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount16` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped16` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id17` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount17` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped17` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id18` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount18` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped18` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id19` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount19` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped19` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id20` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount20` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped20` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id21` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount21` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped21` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id22` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount22` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped22` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id23` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount23` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped23` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id24` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount24` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped24` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id25` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount25` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped25` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
       `item_id26` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount26` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped26` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id27` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount27` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped27` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id28` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount28` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped28` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id29` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount29` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped29` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id30` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount30` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped30` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id31` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount31` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped31` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id32` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount32` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped32` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id33` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount33` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped33` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id34` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount34` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped34` tinyint(1) UNSIGNED NULL DEFAULT '0',
    
      `item_id35` smallint(5) UNSIGNED NULL DEFAULT '0',
      `amount35` smallint(5) UNSIGNED NULL DEFAULT '0',
      `is_equipped35` tinyint(1) UNSIGNED NULL DEFAULT '0') ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;


--
-- Estructura de tabla para la tabla `pet`
--

CREATE TABLE `pet` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `pet1` smallint(5) UNSIGNED DEFAULT NULL,
  `pet2` smallint(5) UNSIGNED DEFAULT NULL,
  `pet3` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `punishment`
--

CREATE TABLE `punishment` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `number` tinyint(3) UNSIGNED NOT NULL,
  `reason` varchar(255) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `skillpoint`
--

CREATE TABLE `skillpoint` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `sk1` tinyint(3) UNSIGNED NOT NULL,
  `exp1` int(10) UNSIGNED NOT NULL,
  `elu1` int(10) UNSIGNED NOT NULL,
  
  `sk2` tinyint(3) UNSIGNED NOT NULL,
  `exp2` int(10) UNSIGNED NOT NULL,
  `elu2` int(10) UNSIGNED NOT NULL,
  
  `sk3` tinyint(3) UNSIGNED NOT NULL,
  `exp3` int(10) UNSIGNED NOT NULL,
  `elu3` int(10) UNSIGNED NOT NULL,
  
  `sk4` tinyint(3) UNSIGNED NOT NULL,
  `exp4` int(10) UNSIGNED NOT NULL,
  `elu4` int(10) UNSIGNED NOT NULL,
  
  `sk5` tinyint(3) UNSIGNED NOT NULL,
  `exp5` int(10) UNSIGNED NOT NULL,
  `elu5` int(10) UNSIGNED NOT NULL,
  
  `sk6` tinyint(3) UNSIGNED NOT NULL,
  `exp6` int(10) UNSIGNED NOT NULL,
  `elu6` int(10) UNSIGNED NOT NULL,
  
  `sk7` tinyint(3) UNSIGNED NOT NULL,
  `exp7` int(10) UNSIGNED NOT NULL,
  `elu7` int(10) UNSIGNED NOT NULL,
  
  `sk8` tinyint(3) UNSIGNED NOT NULL,
  `exp8` int(10) UNSIGNED NOT NULL,
  `elu8` int(10) UNSIGNED NOT NULL,
  
  `sk9` tinyint(3) UNSIGNED NOT NULL,
  `exp9` int(10) UNSIGNED NOT NULL,
  `elu9` int(10) UNSIGNED NOT NULL,
  
  `sk10` tinyint(3) UNSIGNED NOT NULL,
  `exp10` int(10) UNSIGNED NOT NULL,
  `elu10` int(10) UNSIGNED NOT NULL,
  
  `sk11` tinyint(3) UNSIGNED NOT NULL,
  `exp11` int(10) UNSIGNED NOT NULL,
  `elu11` int(10) UNSIGNED NOT NULL,
  
  `sk12` tinyint(3) UNSIGNED NOT NULL,
  `exp12` int(10) UNSIGNED NOT NULL,
  `elu12` int(10) UNSIGNED NOT NULL,
  
  `sk13` tinyint(3) UNSIGNED NOT NULL,
  `exp13` int(10) UNSIGNED NOT NULL,
  `elu13` int(10) UNSIGNED NOT NULL,
  
  `sk14` tinyint(3) UNSIGNED NOT NULL,
  `exp14` int(10) UNSIGNED NOT NULL,
  `elu14` int(10) UNSIGNED NOT NULL,
  
  `sk15` tinyint(3) UNSIGNED NOT NULL,
  `exp15` int(10) UNSIGNED NOT NULL,
  `elu15` int(10) UNSIGNED NOT NULL,
  
  `sk16` tinyint(3) UNSIGNED NOT NULL,
  `exp16` int(10) UNSIGNED NOT NULL,
  `elu16` int(10) UNSIGNED NOT NULL,
  
  `sk17` tinyint(3) UNSIGNED NOT NULL,
  `exp17` int(10) UNSIGNED NOT NULL,
  `elu17` int(10) UNSIGNED NOT NULL,
  
  `sk18` tinyint(3) UNSIGNED NOT NULL,
  `exp18` int(10) UNSIGNED NOT NULL,
  `elu18` int(10) UNSIGNED NOT NULL,
  
  `sk19` tinyint(3) UNSIGNED NOT NULL,
  `exp19` int(10) UNSIGNED NOT NULL,
  `elu19` int(10) UNSIGNED NOT NULL,
  
  `sk20` tinyint(3) UNSIGNED NOT NULL,
  `exp20` int(10) UNSIGNED NOT NULL,
  `elu20` int(10) UNSIGNED NOT NULL,
  
  `sk21` tinyint(3) UNSIGNED NOT NULL,
  `exp21` int(10) UNSIGNED NOT NULL,
  `elu21` int(10) UNSIGNED NOT NULL,
  
  `sk22` tinyint(3) UNSIGNED NOT NULL,
  `exp22` int(10) UNSIGNED NOT NULL,
  `elu22` int(10) UNSIGNED NOT NULL,
  
  `sk23` tinyint(3) UNSIGNED NOT NULL,
  `exp23` int(10) UNSIGNED NOT NULL,
  `elu23` int(10) UNSIGNED NOT NULL,
  
  `sk24` tinyint(3) UNSIGNED NOT NULL,
  `exp24` int(10) UNSIGNED NOT NULL,
  `elu24` int(10) UNSIGNED NOT NULL,
  
  `sk25` tinyint(3) UNSIGNED NOT NULL,
  `exp25` int(10) UNSIGNED NOT NULL,
  `elu25` int(10) UNSIGNED NOT NULL,
  
  `sk26` tinyint(3) UNSIGNED NOT NULL,
  `exp26` int(10) UNSIGNED NOT NULL,
  `elu26` int(10) UNSIGNED NOT NULL,
  
  `sk27` tinyint(3) UNSIGNED NOT NULL,
  `exp27` int(10) UNSIGNED NOT NULL,
  `elu27` int(10) UNSIGNED NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `spell`
--

CREATE TABLE `spell` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `spell_id1` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id2` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id3` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id4` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id5` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id6` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id7` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id8` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id9` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id10` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id11` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id12` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id13` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id14` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id15` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id16` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id17` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id18` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id19` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id20` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id21` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id22` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id23` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id24` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id25` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id26` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id27` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id28` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id29` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id30` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id31` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id32` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id33` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id34` smallint(5) UNSIGNED DEFAULT 0,
  `spell_id35` smallint(5) UNSIGNED DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `familiar`
--

CREATE TABLE `familiar` (
  `user_id` mediumint(8) UNSIGNED NOT NULL,
  `nombre` varchar(30) NOT NULL,
  `level` smallint(5) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL,
  `tipo` int(4) UNSIGNED NOT NULL,
  `min_hp` smallint(5) UNSIGNED NOT NULL,
  `max_hp` smallint(5) UNSIGNED NOT NULL,
  `min_hit` smallint(5) UNSIGNED NOT NULL,
  `max_hit` smallint(5) UNSIGNED NOT NULL,
  `h_id1` smallint(5) UNSIGNED DEFAULT 0,
  `h_id2` smallint(5) UNSIGNED DEFAULT 0,
  `h_id3` smallint(5) UNSIGNED DEFAULT 0,
  `h_id4` smallint(5) UNSIGNED DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `personaje`
--

CREATE TABLE `personaje` (
  `id` mediumint(8) UNSIGNED NOT NULL,
  `cuenta_id` mediumint(8) UNSIGNED NOT NULL,
  `deleted` tinyint(1) NOT NULL DEFAULT 0,
  `name` varchar(30) NOT NULL,
  `level` smallint(5) UNSIGNED NOT NULL,
  `exp` int(10) UNSIGNED NOT NULL,
  `free_skillpoints` int(10) UNSIGNED NOT NULL,
  `assigned_skillpoints` int(10) UNSIGNED NOT NULL,
  `elu` int(10) UNSIGNED NOT NULL,
  `genre_id` tinyint(3) UNSIGNED NOT NULL,
  `race_id` tinyint(3) UNSIGNED NOT NULL,
  `class_id` tinyint(3) UNSIGNED NOT NULL,
  `home_id` tinyint(3) UNSIGNED NOT NULL,
  `description` varchar(255) DEFAULT NULL,
  `gold` int(10) UNSIGNED NOT NULL,
  `bank_gold` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `pet_amount` tinyint(3) UNSIGNED NOT NULL DEFAULT 0,
  `votes_amount` smallint(5) UNSIGNED DEFAULT 0,
  `pos_map` smallint(5) UNSIGNED NOT NULL,
  `pos_x` tinyint(3) UNSIGNED NOT NULL,
  `pos_y` tinyint(3) UNSIGNED NOT NULL,
  `last_map` tinyint(3) UNSIGNED NOT NULL DEFAULT 1,
  `body_id` smallint(5) UNSIGNED NOT NULL,
  `head_id` smallint(5) UNSIGNED NOT NULL,
  `weapon_id` smallint(5) UNSIGNED NOT NULL,
  `helmet_id` smallint(5) UNSIGNED NOT NULL,
  `shield_id` smallint(5) UNSIGNED NOT NULL,
  `aura_id` int(24) DEFAULT 0,
  `aura_color` int(24) DEFAULT 0,
  `heading` tinyint(3) UNSIGNED NOT NULL DEFAULT 3,
  `items_amount` tinyint(3) UNSIGNED NOT NULL,
  `slot_armour` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_weapon` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_nudillos` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_helmet` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_shield` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ammo` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ship` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_ring` tinyint(3) UNSIGNED DEFAULT NULL,
  `slot_bag` tinyint(3) UNSIGNED DEFAULT NULL,
  `min_hp` smallint(5) UNSIGNED NOT NULL,
  `max_hp` smallint(5) UNSIGNED NOT NULL,
  `min_man` smallint(5) UNSIGNED NOT NULL,
  `max_man` smallint(5) UNSIGNED NOT NULL,
  `min_sta` smallint(5) UNSIGNED NOT NULL,
  `max_sta` smallint(5) UNSIGNED NOT NULL,
  `min_ham` smallint(5) UNSIGNED NOT NULL,
  `max_ham` smallint(5) UNSIGNED NOT NULL,
  `min_sed` smallint(5) UNSIGNED NOT NULL,
  `max_sed` smallint(5) UNSIGNED NOT NULL,
  `min_hit` smallint(5) UNSIGNED NOT NULL,
  `max_hit` smallint(5) UNSIGNED NOT NULL,
  `killed_npcs` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `killed` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `rep_asesino` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_bandido` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_burgues` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_ladron` mediumint(8) UNSIGNED NOT NULL DEFAULT 0,
  `rep_noble` mediumint(8) UNSIGNED NOT NULL,
  `rep_plebe` mediumint(8) UNSIGNED NOT NULL,
  `rep_average` mediumint(9) NOT NULL,
  `is_naked` tinyint(1) NOT NULL DEFAULT 0,
  `is_poisoned` tinyint(1) NOT NULL DEFAULT 0,
  `is_incinerado` tinyint(1) DEFAULT 0,
  `is_hidden` tinyint(1) NOT NULL DEFAULT 0,
  `is_hungry` tinyint(1) NOT NULL DEFAULT 0,
  `is_thirsty` tinyint(1) NOT NULL DEFAULT 0,
  `is_ban` tinyint(1) NOT NULL DEFAULT 0,
  `is_dead` tinyint(1) NOT NULL DEFAULT 0,
  `is_sailing` tinyint(1) NOT NULL DEFAULT 0,
  `is_paralyzed` tinyint(1) NOT NULL DEFAULT 0,
  `is_logged` tinyint(1) NOT NULL DEFAULT 0,
  `counter_pena` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `counter_connected` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `counter_training` int(10) UNSIGNED NOT NULL DEFAULT 0,
  `pertenece_consejo_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_consejo_caos` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_real` tinyint(1) NOT NULL DEFAULT 0,
  `pertenece_caos` tinyint(1) NOT NULL DEFAULT 0,
  `ciudadanos_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `criminales_matados` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `recibio_armadura_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_armadura_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_real` tinyint(1) NOT NULL DEFAULT 0,
  `recibio_exp_caos` tinyint(1) NOT NULL DEFAULT 0,
  `recompensas_real` tinyint(3) UNSIGNED DEFAULT 0,
  `recompensas_caos` tinyint(3) UNSIGNED DEFAULT 0,
  `reenlistadas` smallint(5) UNSIGNED NOT NULL DEFAULT 0,
  `fecha_ingreso` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `nivel_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `matados_ingreso` smallint(5) UNSIGNED DEFAULT NULL,
  `siguiente_recompensa` smallint(5) UNSIGNED DEFAULT NULL,
  `guild_index` smallint(5) UNSIGNED DEFAULT 0,
  `guild_aspirant_index` smallint(5) UNSIGNED DEFAULT NULL,
  `guild_member_history` varchar(1024) DEFAULT NULL,
  `guild_requests_history` varchar(1024) DEFAULT NULL,
  `guild_rejected_because` varchar(255) DEFAULT NULL,
  `is_global` tinyint(1) DEFAULT 1,
  `modocombate` tinyint(4) DEFAULT 0,
  `seguro` tinyint(1) DEFAULT 0,
  `pareja` varchar(30) DEFAULT ''
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `macros`
--

CREATE TABLE `macros` (
	  `user_id` mediumint(8) UNSIGNED NOT NULL,
      `tipoaccion1` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell1` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv1` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command1` varchar(24),
	  
	  `tipoaccion2` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell2` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv2` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command2` varchar(24),
	  
	  `tipoaccion3` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell3` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv3` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command3` varchar(24),
	  
	  `tipoaccion4` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell4` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv4` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command4` varchar(24),
	  
	  `tipoaccion5` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell5` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv5` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command5` varchar(24),
	  
	  `tipoaccion6` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell6` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv6` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command6` varchar(24),
	  
	  `tipoaccion7` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell7` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv7` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command7` varchar(24),
	  
	  `tipoaccion8` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell8` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv8` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command8` varchar(24),
	  
	  `tipoaccion9` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell9` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv9` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command9` varchar(24),
	  
	  `tipoaccion10` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell10` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv10` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command10` varchar(24),
	  
	  `tipoaccion11` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell11` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv11` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command11` varchar(24),
	  
	  `tipoaccion12` tinyint(1) UNSIGNED NULL DEFAULT '0',
      `spell12` smallint(5) UNSIGNED NULL DEFAULT '0',
      `inv12` smallint(5) UNSIGNED NULL DEFAULT '0',
	  `command12` varchar(24)
	  
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4; 

--
-- Indices de la tabla `atributos`
--
ALTER TABLE `atributos`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `pet`
--
ALTER TABLE `pet`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD PRIMARY KEY (`user_id`,`number`);

--
-- Indices de la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `spell`
--
ALTER TABLE `spell`
  ADD PRIMARY KEY (`user_id`);

--
-- Indices de la tabla `familiar`
--
ALTER TABLE `familiar`
  ADD PRIMARY KEY (`user_id`);
--
-- Indices de la tabla `personaje`
--
ALTER TABLE `personaje`
  ADD PRIMARY KEY (`id`),
  ADD KEY `name` (`name`);
  
--
-- Indices de la tabla `macros`
--
ALTER TABLE `macros`
  ADD PRIMARY KEY (`user_id`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `personaje`
--
ALTER TABLE `personaje`
  MODIFY `id` mediumint(8) UNSIGNED NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=15;
--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `atributos`
--
ALTER TABLE `atributos`
  ADD CONSTRAINT `fk_atributos_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `banco_items`
--
ALTER TABLE `banco_items`
  ADD CONSTRAINT `fk_bank_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `inventario_items`
--
ALTER TABLE `inventario_items`
  ADD CONSTRAINT `fk_inventory_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `pet`
--
ALTER TABLE `pet`
  ADD CONSTRAINT `fk_pet_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `punishment`
--
ALTER TABLE `punishment`
  ADD CONSTRAINT `fk_punishment_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `skillpoint`
--
ALTER TABLE `skillpoint`
  ADD CONSTRAINT `fk_skillpoint_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

--
-- Filtros para la tabla `spell`
--
ALTER TABLE `spell`
  ADD CONSTRAINT `fk_spell_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);
  
 --
-- Filtros para la tabla `familiar`
--
ALTER TABLE `familiar`
  ADD CONSTRAINT `fk_familiar_user` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);
  
--
-- Filtros para la tabla `macros`
--
ALTER TABLE `macros`
  ADD CONSTRAINT `fk_user_macros` FOREIGN KEY (`user_id`) REFERENCES `personaje` (`id`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;