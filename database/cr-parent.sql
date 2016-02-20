-- Table: Parent
CREATE TABLE parent(
	parent_id INT NOT NULL AUTO_INCREMENT,
	parent_firstname VARCHAR(20) NOT NULL,
	parent_lastname VARCHAR(20) NOT NULL,
	parent_age INT(2) NOT NULL,
	parent_zip CHAR(5) NOT NULL,
	parent_language VARCHAR(20),
	PRIMARY KEY (parent_id)
	);