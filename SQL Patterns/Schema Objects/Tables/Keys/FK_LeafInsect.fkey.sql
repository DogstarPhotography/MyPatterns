﻿ALTER TABLE [dbo].[Insects]
    ADD CONSTRAINT [FK_LeafInsect] FOREIGN KEY ([LeafIndex]) REFERENCES [dbo].[Leaves] ([Index]) ON DELETE NO ACTION ON UPDATE NO ACTION;

