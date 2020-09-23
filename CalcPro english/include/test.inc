function test($myvar) 
	define $i = 0, $j = 1;
	if ($myvar <= 0)
		setError($myvar must be > 0);	
	else
		
		while($i<$myvar)
			$i = $i + 1;
			$j = $j * $i;
		loop
		
		return $j;
	end if

end function