#include <iostream>

extern "C" __declspec(dllexport) void HelloWorld();

void HelloWorld()
{
	std::cout << "Hello World from Cpp";
}