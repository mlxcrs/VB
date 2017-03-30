#include<stdio.h>;
#include<stdlib.h>;
#include<conio.h>;

int p =0;
char f[100];
void e0();
void e1();
void e2();
void aceita();
void rejeita();

int main(){
	p=0;
	printf("Palavra?");
	gets(f);
	e0();
	return 0;
}
void e0(){
	if(f[p]=='a'){
		p++;
		e1();
	}
	else{
		rejeita();
	}
}
void e1(){
	if(f[p]=='b'){
		p++;
		e0();
	}
	else if(f[p]=='c'){
		p++;
		e2();
	}
	else if(f[p]=='a'){
		p++;
		e1();
	}
	else{
		rejeita();
	}
}
void e2(){
	if(f[p]=='a'){
		p++;
		e1();
	}
	else if(f[p]==0){
		aceita();
	}
	else{
		rejeita();
	}
}
void aceita(){
	printf("Aceita \n");
	getch();
	exit(0);
}
void rejeita(){
	printf("Rejeita \n");
	getch();
	exit(0);
}
