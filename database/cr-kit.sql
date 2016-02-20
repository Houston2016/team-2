-- Table: Kit
CREATE TABLE kit(
	kit_id INT NOT NULL AUTO_INCREMENT,
	kit_name VARCHAR(30) NOT NULL,
	kit_book VARCHAR(30) NOT NULL,
	activity_id INT NOT NULL,
	category_id INT NOT NULL,
	kit_path VARCHAR(50),
	PRIMARY KEY (kit_id)
	);