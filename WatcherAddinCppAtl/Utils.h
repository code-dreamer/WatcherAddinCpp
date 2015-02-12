#pragma once

#define NO_COPY_ASSIGN(ClassName)				\
	private:										\
		ClassName(const ClassName &);				\
		ClassName(const ClassName &&);				\
		ClassName &operator=(const ClassName &);	\
		ClassName &operator=(const ClassName &&);


#ifdef UNUSED
#error "UNUSED already defined"	
#else
#define UNUSED(expr) /* expr */
#endif

