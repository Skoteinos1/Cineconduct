# Cineconduct

Ticketing software for cinemas.

This was my first large project. I wrote it roughly one year after I finished my studies. I only wrote in Turbo Pascal before. So pretty much no previous experience. I believe some parts could be done better and probably you will not find enough coments inside. And if you do, they will be in Slovakian language. This code was really not meant for anyone else to read. So please do not judge me based on this.

Software is divided into 2 parts. "Manager" and "Pokladna" (Cashmachine). First Manager enters new movies into database then Cashier sells tickets to customers for those movies. At the end fo the day Cashier prints daily summaries. And at end of the month Manager prints result for whole month.

Every Cinema has different map of seats, I had to write code to generate .ini file with seat map for every cinema. Otherwise I would have to compile updates for every single cinema separately.
