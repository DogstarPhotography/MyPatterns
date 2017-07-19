ALTER TABLE [dbo].[Insects]
    ADD CONSTRAINT [FK_BranchInsect] FOREIGN KEY ([BranchIndex]) REFERENCES [dbo].[Branches] ([Index]) ON DELETE NO ACTION ON UPDATE NO ACTION;

