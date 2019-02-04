CREATE FUNCTION [dbo].[CONSIZE](@bin AS image)
RETURNS int
AS
BEGIN
	DECLARE @pos AS int;
	SET @pos = 1;
	DECLARE @i AS int;
	SET @i = 0;
	DECLARE @ret AS int;
	SET @ret = 0;
	DECLARE @off AS int;
	IF SUBSTRING(@bin, 1, 2) = 0x07FD
		BEGIN
			SET @pos = @pos + 2;
			WHILE SUBSTRING(@bin, @pos, 1) <> 0xFF
				BEGIN
					IF SUBSTRING(@bin, @pos, 1) = 0x00 --STRING
						BEGIN
							SET @pos = @pos + 1;
							SET @off = 0;
							WHILE SUBSTRING(@bin, @pos + @off, 2) <> 0x0000
								BEGIN
									SET @off = @off + 2;
								END
							SET @pos = @pos + @off + 2;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x01 --INT
						BEGIN
							SET @pos = @pos + 5;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x02 --REAL
						BEGIN
							SET @pos = @pos + 11;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x03 --DATE
						BEGIN
							SET @pos = @pos + 4;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x04 --ENUM
						BEGIN
							SET @pos = @pos + 4;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x06 --DATETIME
						BEGIN
							SET @pos = @pos + 13;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x07 --CONTAINER
						BEGIN
							SET @pos = @pos + 1 + dbo.CONSIZE(SUBSTRING(@bin, @pos + 1, DATALENGTH(@bin) - @pos));
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x2D --GUID
						BEGIN
							SET @pos = @pos + 17;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x30 --BLOB
						BEGIN
							SET @pos = @pos + 1;
							SET @off = CAST(CAST(REVERSE(SUBSTRING(@bin, @pos, 4)) AS binary(4)) AS int);
							SET @pos = @pos + 4 + @off;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x31 --INT64
						BEGIN
							SET @pos = @pos + 9;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0xFC --ENUMLABEL
						BEGIN
							SET @pos = @pos + 2;
							SET @off = 0;
							WHILE SUBSTRING(@bin, @pos + @off, 2) <> 0x0000
								BEGIN
									SET @off = @off + 2;
								END
							SET @pos = @pos + @off + 2;
						END
				END
			SET @ret = @pos;
		END
	ELSE
		SET @ret = 0;
	RETURN @ret
END