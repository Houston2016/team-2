-- Table: Library
CREATE TABLE library(
	library_id INT NOT NULL AUTO_INCREMENT,
	library_name VARCHAR(20) NOT NULL,
	library_address VARCHAR(75),
	library_zip CHAR(5) NOT NULL,
	PRIMARY KEY (library_id)
	);