//////////////////////////////////////////////////////////////////////////
//									//
//	datoteka: 	math.inc					//
//									//
//	warning:	please do not change this file, if you		//
//			write wrong commands then program will		//
//			fail on start					//
//									//
//									//
//////////////////////////////////////////////////////////////////////////
//
function sec($x)
	return 1/cos($x);
end function
//
function cosec($x)
	return 1/sin($x);
end function
//
function cotan($x)
	return 1/tan($x);
end function
//
function arcsin($x)
	return atan($x / sqr(-$x * $x + 1));
end function
//
function arccos($x)
	return atan(-$x / sqr(-$x * $x + 1)) + 2 * atan(1);
end function
//
function arcsec($x)
	return atan($x / sqr($x * $x- 1)) + sgn(($x) - 1) * (2 * atan(1));
end function
//
function arccosec($x)
	return atan($x / sqr($x * $x - 1)) + (sgn($x) - 1) * (2 * atan(1));
end function
//
function arccotan($x)
	return atan($x) + 2 * atan(1);
end function
//
function sinh($x)
	return  (exp($x) - exp(-$x)) / 2 ;
end function
//
function cosh($x)
	return  (exp($x) + exp(-$x)) / 2;
end function
//
function tanh($x)
	return  (exp($x) - exp(-$x)) / (exp($x) + exp(-$x));
end function
//
function sech($x)
	return  2 / (exp($x) + exp(-$x));
end function
//
function cosech($x)
	return  2 / (exp($x) - exp(-$x));
end function
//
function cotanh($x)
	return  (exp($x) + exp(-$x)) / (exp($x) - exp(-$x));
end function
//
function arcsinh($x)
	return  ln($x + sqr($x * $x + 1));
end function
//
function arccosh($x)
	return  ln($x + sqr($x * $x - 1));
end function
//
function atanh($x)
	return  ln((1 + $x) / (1 - $x)) / 2;
end function
//
function arcsech($x)
	return ln((sqr(-$x * $x + 1) + 1) / $x);
end function
//
function arccosech($x)
	return ln((sgn($x) * sqr($x * $x + 1) + 1) / $x);
end function
//
function arccotanh($x)
	return ln(($x + 1) / ($x - 1)) / 2;
end function