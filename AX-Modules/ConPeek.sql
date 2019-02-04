CREATE FUNCTION [dbo].[CONPEEK](@bin AS varbinary(8000), @idx AS int)
RETURNS sql_variant
AS
BEGIN
	DECLARE @pos AS int;
	SET @pos = 1;
	DECLARE @i AS int;
	SET @i = 1;
	DECLARE @off AS int;
	DECLARE @ret AS sql_variant;
	IF SUBSTRING(@bin, @pos, 2) = 0x07FD
		BEGIN
			SET @pos = @pos + 2;
			WHILE @idx > 0 AND SUBSTRING(@bin, @pos, 1) <> 0xFF
				BEGIN
					IF SUBSTRING(@bin, @pos, 1) = 0x00 --STRING
						BEGIN
							SET @pos = @pos + 1;
							SET @off = 0;
							SET @ret = '';
							WHILE SUBSTRING(@bin, @pos + @off, 2) <> 0x0000
								BEGIN
									SET @ret = CAST(@ret AS varchar(8000)) + 
										CHAR(CAST(REVERSE(SUBSTRING(@bin, @pos + @off, 2)) AS binary(2)))
									SET @off = @off + 2;
								END
							SET @pos = @pos + @off + 2;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x01 --INT
						BEGIN
							SET @pos = @pos + 1;
							SET @ret = CAST(REVERSE(SUBSTRING(@bin, @pos, 4)) AS binary(4));
							SET @ret = CAST(@ret AS int);
							SET @pos = @pos + 4;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x02 --REAL
						BEGIN
							SET @pos = @pos + 1;
							DECLARE @temp as binary(8);
							SET @temp = CAST(REVERSE(SUBSTRING(@bin, @pos + 2, 8)) AS binary(8));
							DECLARE @val bigint;
							SET @val = 0;
							DECLARE @dec AS int;
							SET @off = 1;
							WHILE (@off <= 8)
								BEGIN
									SET @val = (@val * 100) +
										(CAST(SUBSTRING(@temp, @off, 1) AS int) / 0x10 * 10) +
										(CAST(SUBSTRING(@temp, @off, 1) AS int) % 0x10);
									SET @off = @off + 1;
								END
							WHILE @val <> 0 AND @val % 10 = 0
								SET @val = @val / 10;
							SET @dec = (CAST(SUBSTRING(@bin, @pos, 1) AS int) + 0x80) % 0x100;
							SET @ret = CAST(@val AS real);
							WHILE @dec >= LEN(CAST(@val AS varchar)) + 0x80
								BEGIN
									SET @ret = CAST(@ret AS real) * 10.0;
									SET @dec = @dec - 1;
								END
							SET @dec = (CAST(SUBSTRING(@bin, @pos, 1) AS int) + 0x80) % 0x100 + 1;
							WHILE @dec < LEN(CAST(@val AS varchar)) + 0x80
								BEGIN
									SET @ret = CAST(@ret AS real) / 10.0;
									SET @dec = @dec + 1;
								END
							IF SUBSTRING(@bin, @pos + 1, 1) = 0x80
								SET @ret = 0 - CAST(@ret AS real);
							SET @pos = @pos + 10;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x03 --DATE
						BEGIN
							SET @pos = @pos + 1;
							DECLARE @year char(4);
							DECLARE @month char(2);
							DECLARE @day char(2);
							SET @year = SUBSTRING(@bin, @pos, 1) + 1900;
							SET @month = SUBSTRING(@bin, @pos + 1, 1) + 1;
							SET @day = SUBSTRING(@bin, @pos + 2, 1) + 1;
							IF LEN(@month) < 2
								SET @month = '0' + @month;
							IF LEN(@day) < 2
								SET @day = '0' + @day;
							SET @ret = CAST(@year + '-' + @month + '-' + @day AS date);
							SET @pos = @pos + 3;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x04 --ENUM
						BEGIN
							SET @pos = @pos + 1;
							SET @ret = CAST(SUBSTRING(@bin, @pos, 1) AS int);
							SET @pos = @pos + 3;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x06 --DATETIME
						BEGIN
							SET @pos = @pos + 1;
							DECLARE @year2 char(4);
							DECLARE @month2 char(2);
							DECLARE @day2 char(2);
							DECLARE @hour char(2);
							DECLARE @min char(2);
							DECLARE @sec char(2);
							SET @year2 = SUBSTRING(@bin, @pos, 1) + 1900;
							SET @month2 = SUBSTRING(@bin, @pos + 1, 1) + 1;
							SET @day2 = SUBSTRING(@bin, @pos + 2, 1) + 1;
							SET @hour = SUBSTRING(@bin, @pos + 3, 1) + 0;
							SET @min = SUBSTRING(@bin, @pos + 4, 1) + 0;
							SET @sec = SUBSTRING(@bin, @pos + 5, 1) + 0;
							IF LEN(@month2) < 2
								SET @month2 = '0' + @month2;
							IF LEN(@day2) < 2
								SET @day2 = '0' + @day2;
							IF LEN(@hour) < 2
								SET @hour = '0' + @hour;
							IF LEN(@min) < 2
								SET @min = '0' + @min;
							IF LEN(@sec) < 2
								SET @sec = '0' + @sec;
							SET @ret = CAST(@year2 + '-' + @month2 + '-' + @day2 + ' ' + @hour + ':' + @min + ':' + @sec AS datetime);
							SET @pos = @pos + 12;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x07 --CONTAINER
						BEGIN
							SET @pos = @pos + 1;
							DECLARE @len int;
							SET @len = DATALENGTH(@bin) - (@pos - 1);
							SET @len = dbo.CONSIZE(SUBSTRING(@bin, @pos, @len));
							SET @ret = SUBSTRING(@bin, @pos, @len);
							SET @pos = @pos + @len;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x2D --GUID
						BEGIN
							SET @pos = @pos + 1;
							SET @off = 0;
							SET @ret = CAST(
								REVERSE(SUBSTRING(@bin, @pos, 4)) +
								REVERSE(SUBSTRING(@bin, @pos + 4, 4)) +
								REVERSE(SUBSTRING(@bin, @pos + 8, 4)) +
								REVERSE(SUBSTRING(@bin, @pos + 12, 4)
							) AS binary(16));
							SET @ret = CAST(@ret AS uniqueidentifier);
							SET @pos = @pos + 16;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x30 --BLOB
						BEGIN
							SET @pos = @pos + 1;
							SET @off = CAST(CAST(REVERSE(SUBSTRING(@bin, @pos, 4)) AS binary(4)) AS int);
							SET @pos = @pos + 4;
							SET @ret = CAST(SUBSTRING(@bin, @pos, @off) AS varbinary(8000));
							SET @pos = @pos + @off;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0x31 --INT64
						BEGIN
							SET @pos = @pos + 1;
							SET @ret = CAST(CAST(REVERSE(SUBSTRING(@bin, @pos, 8)) AS binary(8)) AS bigint);
							SET @pos = @pos + 8;
						END
					ELSE IF SUBSTRING(@bin, @pos, 1) = 0xFC --ENUMLABEL
						BEGIN
							SET @pos = @pos + 1;
							DECLARE @value int;
							DECLARE @name varchar(40);
							SET @value = SUBSTRING(@bin, @pos, 1);
							SET @pos = @pos + 1;
							SET @off = 0;
							SET @name = '';
							WHILE SUBSTRING(@bin, @pos + @off, 2) <> 0x0000
								BEGIN
									SET @name = CAST(@name AS varchar(40)) + 
										CHAR(CAST(REVERSE(SUBSTRING(@bin, @pos + @off, 2)) AS binary(2)))
									SET @off = @off + 2;
								END
							SET @ret = CAST(@name + ':' + CAST(@value as varchar(3)) as varchar(44));
							SET @pos = @pos + @off + 2;
						END
					ELSE
						SET @ret = NULL;
					IF @i = @idx
						RETURN @ret
					ELSE
						BEGIN
							SET @i = @i + 1;
							SET @ret = NULL;
						END
				END
			RETURN NULL;
		END
	RETURN @ret
END