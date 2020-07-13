<?php

$servername = "localhost";
$database = "localhostdb";
$username = "lectura";
$password = "lectura2";

//  Create a new connection to the MySQL database using PDO
$conn = new mysqli($servername, $username, $password, $database);
// Check connection
if ($conn->connect_error) {
   die("Connection failed: " . $conn->connect_error);
}