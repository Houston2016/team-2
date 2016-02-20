-- Table: Evaluation
CREATE TABLE evaluation(
	evaluation_id INT NOT NULL AUTO_INCREMENT,
	evaluation_date DATE NOT NULL,
	parent_id INT NOT NULL,
	kit_id INT NOT NULL,
	evaluation_enjoyed INT(1) NOT NULL,
	evaluation_comment VARCHAR(50),
	PRIMARY KEY (evaluation_id)
	);