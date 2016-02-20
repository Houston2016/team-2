-- Table: Kit
ALTER TABLE kit
ADD FOREIGN KEY (activity_id)
REFERENCES activity(activity_id),
ADD FOREIGN KEY (category_id)
REFERENCES category(category_id)