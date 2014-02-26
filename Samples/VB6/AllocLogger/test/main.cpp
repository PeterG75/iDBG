#include <stdio.h>
#include <windows.h>
#include <conio.h>

void main(void){


	for(int i=0; i < 20 ; i++){

		char* b1 = (char*)GlobalAlloc(0,200);
		char* b2 = (char*)GlobalAlloc(0,100);
		char* b3 = (char*)LocalAlloc(0,100);
		char* b4 = (char*)malloc(100);
		char* b5 = (char*)HeapAlloc(GetProcessHeap(),0,120);
		char* b6 = (char*)VirtualAlloc(0,200,MEM_COMMIT, PAGE_EXECUTE_READWRITE );

		printf("b1=%x\n", (int)b1);
		printf("b2=%x\n", (int)b2);
		printf("b3=%x\n", (int)b3);
		printf("b4=%x\n", (int)b4);
		printf("b5=%x\n", (int)b5);
		printf("b6=%x\n", (int)b6);

		sprintf(b1, "this is my string 1 %x (GlobalAlloc)"  , GetTickCount() );
		sprintf(b2, "this is my string 2 %x (GlobalAlloc)"  , GetTickCount() );
		sprintf(b3, "this is my string 3 %x (LocalAlloc)"   , GetTickCount() );
		sprintf(b4, "this is my string 4 %x (malloc)"       , GetTickCount() );
		sprintf(b5, "this is my string 5 %x (HeapAlloc)"    , GetTickCount() );
		sprintf(b6, "this is my string 6 %x (VirtualAlloc)" , GetTickCount() );

		GlobalFree(b1);
		GlobalFree(b2);
		LocalFree(b3);
		free(b4);
		HeapFree(GetProcessHeap(),0,b5);
		VirtualFree(b6,200,MEM_RELEASE);
		
		printf("----------- *yawn* I'm tired... I am going to sleep ----------\n\n");
		Sleep(2500);
	}
	

	printf("Done");
	getch();


}