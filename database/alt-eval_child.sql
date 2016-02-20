-- Table: Eval_Child
ALTER TABLE eval_child
ADD FOREIGN KEY (evaluation_id)
REFERENCES evaluation(evaluation_id),
ADD FOREIGN KEY (child_id)
REFERENCES child(child_id)