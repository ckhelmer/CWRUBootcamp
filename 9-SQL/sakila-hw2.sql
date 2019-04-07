#7h. List the top five genres in gross revenue in descending order
SELECT name, SUM(amount) AS "Gross Revenue"
FROM sakila.category a
JOIN sakila.film_category b
ON a.category_id = b.category_id
JOIN sakila.inventory c
ON b.film_id = c.film_id
JOIN sakila.rental d
ON c.inventory_id = d.inventory_id
JOIN sakila.payment e
ON d.rental_id = e.rental_id
GROUP BY name
ORDER BY SUM(amount) DESC LIMIT 5;

#8a. Create a view for #7h (above).
CREATE VIEW sakila.Top_Five_Genres 
AS SELECT name, SUM(amount) AS "Gross Revenue"
FROM sakila.category a
JOIN sakila.film_category b
ON a.category_id = b.category_id
JOIN sakila.inventory c
ON b.film_id = c.film_id
JOIN sakila.rental d
ON c.inventory_id = d.inventory_id
JOIN sakila.payment e
ON d.rental_id = e.rental_id
GROUP BY name
ORDER BY SUM(amount) DESC LIMIT 5;

#8b. Display the view
SELECT * FROM sakila.Top_Five_Genres;

#8c. Delete the view
DROP VIEW sakila.Top_Five_Genres; 