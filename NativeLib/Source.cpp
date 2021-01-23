//https://www.youtube.com/watch?v=NKIdxJAbr0Q
#include <iostream>

extern "C" __declspec(dllexport) void HelloWorld();

void HelloWorld()
{
	std::cout << "Hello World from Cpp";
}