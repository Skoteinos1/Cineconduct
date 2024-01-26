# Cineconduct

Ticketing software for cinemas.

This was my first large project. I wrote it roughly one year after I finished my studies. I only wrote in Turbo Pascal before. So pretty much no previous experience. I believe some parts could be done better and probably you will not find abundance of comments within the code. If any are present, they would be in the Slovakian language. This code was not meant for anyone else to read. So please do not judge me solely based on this project.

Software is divided into 2 parts. "Manager" and "Pokladna" (Box Office). First the Manager inputs new movies into the database, after which the Cashier sells tickets to customers for these movies. At the end of each day, the Cashier generates daily summaries, while the Manager compiles results for the entire month at the end of each month.

Every cinema has a unique seating map, I had to write code to generate .ini file with seat map for every cinema. Otherwise I would have to compile updates separately for each cinema.

![Picture of Old and New Software](https://github.com/Skoteinos1/Cineconduct/blob/main/Cineconduct_old_vs_new.png)

