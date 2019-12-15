CREATE TABLE PaymentOrder (
  Id INT PRIMARY KEY,
  IncomingDate DATE NOT NULL,
  OutcomingDate DATE NOT NULL,
  Amount DECIMAL(20, 4) NOT NULL
);

INSERT INTO PaymentOrder(Id, IncomingDate, OutcomingDate, Amount) 
VALUES (1, '2019-12-15', '2019-12-16', 123.45)
INSERT INTO PaymentOrder(Id, IncomingDate, OutcomingDate, Amount) 
VALUES (2, '2019-12-17', '2019-12-18', 777.33)
