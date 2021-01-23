// Test_Cpp.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>

int main()
{
    std::cout << "Hello World!\n";
}

extern void __stdcall SetStringArray(SAFEARRAY& safeArray)
{
    if (safeArray.cDims == 1)
    {
        if ((safeArray.fFeatures & FADF_BSTR) == FADF_BSTR)
        {
            BSTR* bstrArray;
            HRESULT hr = SafeArrayAccessData(&safeArray, (void**)&bstrArray);

            long iMin = 0;
            SafeArrayGetLBound(&safeArray, 1, &iMin);
            long iMax = 0;
            SafeArrayGetUBound(&safeArray, 1, &iMax);

            for (long i = iMin; i <= iMax; ++i)
            {
                // Do something here with the data! 
            }
        }
    }
}
// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
