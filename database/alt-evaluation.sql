-- Table: Evaluation
ALTER TABLE evaluation
ADD FOREIGN KEY (parent_id)
REFERENCES parent(parent_id),
ADD FOREIGN KEY (kit_id)
REFERENCES kit(kit_id)