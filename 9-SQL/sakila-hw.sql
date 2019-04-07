#1a. Display the first and last names of all actors
SELECT first_name, last_name FROM sakila.actor;

#1b. Add a new column with first and last name in uppercase.alter.
SELECT concat(first_name, " ", last_name) AS Actor_Name
FROM sakila.actor; 

#2a. Find ID, First and Last name of Joe
SELECT actor_id, first_name, last_name FROM sakila.actor
WHERE first_name = "Joe";

#2b. Find all actors with last name "GEN"
SELECT * FROM sakila.actor
WHERE last_name LIKE "%gen%";

#2c. Find all actors whose last names contain LI. Order the rows by last name then first name.
SELECT last_name, first_name FROM sakila.actor
WHERE last_name LIKE "%li%";

#2d. Using in, display country_id, and country for Afghanistan, Bangladesh, and China
SELECT country_id, country FROM sakila.country
WHERE country in ("Afghanistan", "Bangladesh", "China");

#3a. Add a column to describe actors.
ALTER TABLE sakila.actor
ADD description BLOB;

#3b. Get rid of the column
ALTER TABLE sakila.actor
DROP description;

#4a. List the last names of actors with their counts.
SELECT COUNT(actor_id), last_name FROM sakila.actor
GROUP BY last_name 
ORDER BY COUNT(actor_id) DESC; 

#4b. List last names of actors with number, only if they have 2 actors each
SELECT COUNT(last_name), last_name FROM sakila.actor
GROUP BY last_name
HAVING COUNT(last_name) >= 2 
ORDER BY COUNT(actor_id) DESC; 

#4c. Change Groucho Williams to Harpo Williams
UPDATE sakila.actor 
SET first_name = "HARPO"
WHERE first_name = "GROUCHO" AND last_name = "WILLIAMS";

#4d. Change it back.
UPDATE sakila.actor 
SET first_name = "GROUCHO"
WHERE first_name = "HARPO" AND last_name = "WILLIAMS";

#5a. Locate the schema of the address table. 
SHOW CREATE TABLE sakila.address;

#6a. Use Join to display first and last names of each staff member with address
SELECT first_name, last_name, address
FROM sakila.staff a
JOIN sakila.address b
ON a.address_id = b.address_id; 

#6b. Use Join to display the total amount rung up by each staff member in August 2005.
SELECT first_name, last_name, sum(amount)
FROM sakila.payment a
JOIN sakila.staff b 
ON a.staff_id = b.staff_id
WHERE payment_date LIKE "2005-08%"
GROUP BY first_name, last_name;

#6c.  List each film and the number of actors who are listed for that film. 
SELECT title AS "Film", count(actor_id)  AS "Number of Actors"
FROM sakila.film a
JOIN sakila.film_actor b
ON a.film_id = b.film_id
GROUP BY title;

#6d. How many copies of the film Hunchback Impossible exist in the inventory system?
SELECT title AS "Film" , COUNT(inventory_id) AS "Copies Owned"
FROM sakila.film a
JOIN sakila.inventory b
ON a.film_id = b.film_id
WHERE title = "Hunchback Impossible";

#6e. List amount spent by each customer. List customers alphabetically by last name.
SELECT first_name, last_name, SUM(amount) AS "Total Amount Paid"
FROM sakila.payment a
JOIN sakila.customer b
ON a.customer_id = b.customer_id
GROUP BY last_name
ORDER BY last_name;

#7a. Use subqueries to display films starting with K or Q with English titles.alter
SELECT title, (SELECT name FROM sakila.language WHERE name LIKE "English%") AS LANGUAGE
FROM sakila.film
WHERE title LIKE "Q%" OR title LIKE "K%"; 

#7b. Use subqueries to display all actors who appear in Alone Trip.
SELECT first_name, last_name 
FROM sakila.actor
WHERE actor_id in 
	(SELECT actor_id 
	 FROM sakila.film_actor 
     WHERE film_id IN 
		(SELECT film_id 
        FROM sakila.film 
        WHERE title = "Alone Trip"));

#7c. You want to run an email marketing campaign in Canada, for which you will need the names and email addresses of all Canadian customers. 
SELECT first_name, last_name, email
FROM sakila.customer a
JOIN sakila.address b
ON a.address_id = b.address_id
JOIN sakila.city c
ON b.city_id = c.city_id
JOIN sakila.country d
ON c.country_id = d.country_id
WHERE country = "Canada";

#7d. Identify all movies categorized as family films
SELECT title
FROM sakila.film a
JOIN sakila.film_category b
ON a.film_id = b.film_id
JOIN sakila.category c
ON c.category_id = b.category_id
WHERE c.name = "Family";

#7e. Display the most frequently rented movies in descending order.
SELECT title, COUNT(rental_id) AS "Rentals"
FROM sakila.film a
JOIN sakila.inventory b
ON a.film_id = b.film_id
JOIN sakila.rental c
ON b.inventory_id = c.inventory_id
GROUP BY title
ORDER BY COUNT(rental_id) DESC; 

#7f. Write a query to display how much business, in dollars, each store brought in.
#The easy way:
SELECT * FROM sakila.sales_by_store;

#The hard way:
SELECT address, city, country, SUM(amount) AS "Total Sales"
FROM sakila.store a
JOIN sakila.customer b
ON a.store_id = b.store_id
JOIN sakila.payment c
ON b.customer_id = c.customer_id
JOIN sakila.address d
ON a.address_id = d.address_id
JOIN sakila.city e
ON d.city_id = e.city_id
JOIN sakila.country f
ON e.country_id = f.country_id
GROUP BY a.store_id;

#7g. Write a query to display for each store its store ID, city, and country.
SELECT store_id, city, country 
FROM sakila.store a
JOIN sakila.address b
ON a.address_id = b.address_id
JOIN sakila.city c
ON c.city_id = b.city_id
JOIN sakila.country d
ON c.country_id = d.country_id

#7h. List the top five genres in gross revenue in descending order
SELECT SUM(amount)
FROM sakila.payment a
JOIN sakila.rental b
ON a.rental_id = b.rental_id
JOIN sakila.inventory c
ON b.inventory_id = c.inventory_id
JOIN sakila.film_category d
ON c.film_id = d.film_id
JOIN sakila.category e
ON d.category_id = e.category_id;