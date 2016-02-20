-- Table: Child
ALTER TABLE child
ADD FOREIGN KEY (parent_id)
REFERENCES parent(parent_id)