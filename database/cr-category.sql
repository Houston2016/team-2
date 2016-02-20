-- Table: Category
CREATE TABLE category(
	category_id INT NOT NULL AUTO_INCREMENT,
	category_name VARCHAR(30) NOT NULL,
	category_age1 INT(3) NOT NULL,
	category_age2 INT(3) NOT NULL,
	category_descr VARCHAR(50),
	PRIMARY KEY (category_id)
	);