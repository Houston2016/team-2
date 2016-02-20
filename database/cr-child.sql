-- Table: Child
CREATE TABLE child(
	child_id INT NOT NULL AUTO_INCREMENT,
	child_firstname VARCHAR(20) NOT NULL,
	child_lastname VARCHAR(20) NOT NULL,
	child_age INT(2) NOT NULL,
	parent_id INT NOT NULL,
	PRIMARY KEY (child_id)
	);