-- Table: Libr_Kit
ALTER TABLE libr_kit
ADD FOREIGN KEY (library_id)
REFERENCES library(library_id),
ADD FOREIGN KEY (kit_id)
REFERENCES kit(kit_id)