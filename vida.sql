-- MySQL dump 10.13  Distrib 8.0.18, for Win64 (x86_64)
--
-- Host: 127.0.0.1    Database: vida
-- ------------------------------------------------------
-- Server version	8.4.3

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!50503 SET NAMES utf8mb4 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `pengguna`
--

DROP TABLE IF EXISTS `pengguna`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `pengguna` (
  `id` int NOT NULL AUTO_INCREMENT,
  `nama` varchar(255) NOT NULL,
  `email` varchar(255) NOT NULL,
  `password` varchar(255) NOT NULL,
  `role` enum('EMKL','CS','Bea Cukai','Gate Superintendent') DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `email` (`email`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pengguna`
--

LOCK TABLES `pengguna` WRITE;
/*!40000 ALTER TABLE `pengguna` DISABLE KEYS */;
INSERT INTO `pengguna` VALUES (1,'John Doe','john.doe@example.com','12345','EMKL'),(2,'Jane Smith','jane.smith@example.com','password456','CS'),(3,'Bea Cukai User','beacukai_user@example.com','password_anda','Bea Cukai'),(4,'Gate User 1','gateuser1@example.com','password3','Gate Superintendent');
/*!40000 ALTER TABLE `pengguna` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `permohonan`
--

DROP TABLE IF EXISTS `permohonan`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `permohonan` (
  `id` int NOT NULL AUTO_INCREMENT,
  `id_pengguna` int DEFAULT NULL,
  `nomor_kontainer` varchar(100) NOT NULL,
  `status` enum('Menunggu','Disetujui','Ditolak') NOT NULL DEFAULT 'Menunggu',
  `waktu_pengajuan` datetime DEFAULT CURRENT_TIMESTAMP,
  `catatan` text,
  `nama_perusahaan` varchar(255) DEFAULT NULL,
  `nopol_truck` varchar(255) DEFAULT NULL,
  `file_surat` varchar(200) DEFAULT NULL,
  `kode_ea` varchar(50) DEFAULT NULL,
  `tanggal_pengajuan` varchar(20) DEFAULT NULL,
  `jenis_permohonan` varchar(50) DEFAULT NULL,
  `alasan_permohonan` text,
  `keterangan_tambahan` text,
  `created_at` datetime DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`),
  KEY `id_pengguna` (`id_pengguna`),
  CONSTRAINT `permohonan_ibfk_1` FOREIGN KEY (`id_pengguna`) REFERENCES `pengguna` (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `permohonan`
--

LOCK TABLES `permohonan` WRITE;
/*!40000 ALTER TABLE `permohonan` DISABLE KEYS */;
INSERT INTO `permohonan` VALUES (1,1,'KON123456','Disetujui','2026-04-18 22:32:52',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2026-05-10 20:49:31'),(2,2,'KON789101','Disetujui','2026-04-18 22:32:52','hayo hayo',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2026-05-10 20:49:31'),(3,NULL,'mk2345','Menunggu','2026-04-21 05:50:50',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2026-05-10 20:49:31'),(4,NULL,'kml23','Menunggu','2026-04-21 06:01:51',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2026-05-10 20:49:31'),(5,NULL,'kml23','Menunggu','2026-04-21 06:03:25',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2026-05-10 20:49:31'),(6,NULL,'my78','Menunggu','2026-04-21 06:04:57',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'2026-05-10 20:49:31'),(7,NULL,'mk2345','Menunggu','2026-05-10 20:50:24',NULL,'PT. Seminar muda','w398','surat_permohonan_ea-2025-334_20260510205024.docx','ea-2025-334','2026-05-10','pengeluaran','pengeluaran kontainer','','2026-05-10 13:50:24'),(8,NULL,'mk2345','Menunggu','2026-05-10 20:52:10',NULL,'PT. Seminar muda','w398','surat_permohonan_ea-2025-334_20260510205210.docx','ea-2025-334','2026-05-10','pengeluaran','pengeluaran kontainer','','2026-05-10 13:52:11'),(9,NULL,'mk2345','Menunggu','2026-05-10 20:52:51',NULL,'PT. Seminar muda','w398','surat_permohonan_ea-2025-334_20260510205251.docx','ea-2025-334','2026-05-10','pemasukan','pemasukan','','2026-05-10 13:52:52'),(10,NULL,'mk2345','Menunggu','2026-05-10 21:31:03',NULL,'PT. Seminar muda','w398','surat_permohonan_ea-2025-334_20260510213103.docx','ea-2025-334','2026-05-10','relokasi','relokasi kontainer','','2026-05-10 14:31:04'),(11,1,'my78','Disetujui','2026-05-10 21:40:41','lengkap','PT. Seminar muda','y9087','surat_permohonan_ea-2025-334_20260510214041.docx','ea-2025-334','2026-05-10','relokasi','relokasi','','2026-05-10 14:40:41'),(12,1,'abc768','Disetujui','2026-05-10 22:04:32','lengkap\r\n','PT. Seminar muda','i987','surat_persetujuan_ew-223-098_20260510220516.docx','ew-223-098','2026-05-08','relokasi','relokasi','','2026-05-10 15:04:32'),(13,1,'njk098','Disetujui','2026-05-10 22:20:33','lengkap bray','PT. Seminar muda','y876','surat_persetujuan_tr-678-908_20260510222150.pdf','tr-678-908','2026-05-06','pengeluaran','pengeluaran barang','','2026-05-10 15:20:34'),(14,1,'kl098','Menunggu','2026-05-10 23:13:11',NULL,'PT. Seminar muda','o875','surat_permohonan_kj-09-08_20260510231310.docx','kj-09-08','2026-03-10','pengeluaran','pengeluaran barang','','2026-05-10 16:13:11');
/*!40000 ALTER TABLE `permohonan` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `persetujuan`
--

DROP TABLE IF EXISTS `persetujuan`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!50503 SET character_set_client = utf8mb4 */;
CREATE TABLE `persetujuan` (
  `id` int NOT NULL AUTO_INCREMENT,
  `id_permohonan` int DEFAULT NULL,
  `status_persetujuan` enum('Disetujui','Ditolak') NOT NULL,
  `waktu_persetujuan` datetime DEFAULT CURRENT_TIMESTAMP,
  `komentar` text,
  PRIMARY KEY (`id`),
  KEY `id_permohonan` (`id_permohonan`),
  CONSTRAINT `persetujuan_ibfk_1` FOREIGN KEY (`id_permohonan`) REFERENCES `permohonan` (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `persetujuan`
--

LOCK TABLES `persetujuan` WRITE;
/*!40000 ALTER TABLE `persetujuan` DISABLE KEYS */;
INSERT INTO `persetujuan` VALUES (1,1,'Disetujui','2026-04-18 22:35:24','Permohonan telah memenuhi syarat dan disetujui oleh Bea Cukai'),(2,2,'Ditolak','2026-04-18 22:35:24','Kontainer tidak memenuhi persyaratan dan ditolak.');
/*!40000 ALTER TABLE `persetujuan` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Dumping routines for database 'vida'
--
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2026-05-11  9:29:37
