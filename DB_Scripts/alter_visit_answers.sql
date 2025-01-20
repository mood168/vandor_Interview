ALTER TABLE VisitAnswers
ADD ModifiedBy int NULL
CONSTRAINT FK_VisitAnswers_Users 
FOREIGN KEY (ModifiedBy) REFERENCES Users(UserID) 