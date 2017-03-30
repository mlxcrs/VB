#include<stdio.h>
#include<stdlib.h>
#include<conio.h>

int main(){
	int p;
	char f[100];
	p=0;
	printf("Palavra?\n");
	gets(f);
	goto e0;
	
	e0:
	if(f[p]=='a'){
		p++;
		goto e1;
	}
	
	e1:
	if(f[p]=='b'){
		p++;
		goto e0;
	}
	else if(f[p]=='c'){
		p++;
		goto e2;
	}
	
	e2:
	if(f[p]=='a'){
		p++;
		goto e0;
	}
	else if(f[p]==0){
		goto aceita;
	}
	else{
		goto rejeita;
	}
	
	aceita:
	printf("Aceita\n");
	getch();
	return 0;
	
	rejeita:
	printf("Rejeita\n");
	getch();
	return 0;
}
