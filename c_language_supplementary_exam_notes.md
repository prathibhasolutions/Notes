# C Language Supplementary Exam – Important Questions and Notes

This note is built from the two past question papers and focuses on the most exam-relevant C-language topics only.

---

## 1. Algorithm, Flowchart, and Basic Concepts

### 1.1 What is an algorithm?
An algorithm is a step-by-step procedure to solve a problem.

Important characteristics:
- Finite number of steps
- Clear and unambiguous
- Input and output defined
- Effective and efficient

### 1.2 Example: Algorithm to find the maximum number from a given list
```
Algorithm MaxNumber
Input: n numbers
Output: largest number

1. Read n
2. Read first number as max
3. For i = 2 to n
   a. Read next number
   b. If number > max then max = number
4. Print max
End
```

### 1.3 Flowchart: Check whether a number is perfect or not
A perfect number is a number whose sum of proper divisors equals the number itself.

Example: 6 = 1 + 2 + 3.

Flowchart logic:
- Read number `n`
- Set `sum = 0`
- For `i = 1` to `n/2`
  - If `n % i == 0`, add `i` to `sum`
- If `sum == n`, print `Perfect`
- Else print `Not Perfect`

### 1.4 C program to check perfect number
```c
#include <stdio.h>

int main() {
    int n, i, sum = 0;
    printf("Enter a number: ");
    scanf("%d", &n);

    for (i = 1; i <= n / 2; i++) {
        if (n % i == 0) {
            sum += i;
        }
    }

    if (sum == n)
        printf("%d is a perfect number.\n", n);
    else
        printf("%d is not a perfect number.\n", n);

    return 0;
}
```

---

## 2. RAM vs ROM

### Difference between RAM and ROM
- RAM (Random Access Memory): temporary storage, volatile, used while the program is running.
- ROM (Read Only Memory): permanent storage, non-volatile, stores firmware/boot instructions.

### Quick comparison
- RAM: read/write, volatile, faster, temporary data storage.
- ROM: read only, non-volatile, stores permanent instructions.

---

## 3. Data Types in C

### Basic C data types
- `char` – 1 byte – stores single character
- `int` – usually 4 bytes – stores integer
- `float` – 4 bytes – stores decimal numbers
- `double` – 8 bytes – stores larger decimal values

### Example
```c
#include <stdio.h>

int main() {
    char ch = 'A';
    int age = 21;
    float marks = 78.5f;
    double pi = 3.1415926535;

    printf("char = %c\n", ch);
    printf("int = %d\n", age);
    printf("float = %.2f\n", marks);
    printf("double = %.10f\n", pi);

    return 0;
}
```

### Other important types
- `void` – no value
- `short` / `long` – sized integer types
- `unsigned` – only non-negative values

---

## 4. Control Statements in C

### Control statements include
- `if`, `if-else`, `else if`
- `switch`
- `for`, `while`, `do-while`
- `break`, `continue`, `goto`

### Example: Factorial using loop
```c
#include <stdio.h>

int main() {
    int n, i, fact = 1;

    printf("Enter a number: ");
    scanf("%d", &n);

    for (i = 1; i <= n; i++) {
        fact *= i;
    }

    printf("Factorial of %d = %d\n", n, fact);
    return 0;
}
```

### Example: `switch` statement
```c
#include <stdio.h>

int main() {
    int choice;
    printf("1. C\n2. Java\n3. Python\nChoose: ");
    scanf("%d", &choice);

    switch (choice) {
        case 1:
            printf("C selected\n");
            break;
        case 2:
            printf("Java selected\n");
            break;
        case 3:
            printf("Python selected\n");
            break;
        default:
            printf("Invalid choice\n");
    }

    return 0;
}
```

---

## 5. Functions in C

### Function definition and call
A function is a reusable block of code.

### Example: Fibonacci series using function
```c
#include <stdio.h>

void fibSeries(int n) {
    int a = 0, b = 1, c;

    printf("Fibonacci Series: ");
    for (int i = 0; i < n; i++) {
        printf("%d ", a);
        c = a + b;
        a = b;
        b = c;
    }
    printf("\n");
}

int main() {
    int n;
    printf("Enter how many terms: ");
    scanf("%d", &n);
    fibSeries(n);
    return 0;
}
```

### Important function concepts
- Function prototype
- Call by value
- Call by reference
- Return type
- Recursion

---

## 6. Arrays and Strings

### 6.1 Define an array
An array is a collection of similar data types stored in contiguous memory locations.

### Declaration and initialization
```c
int a[5] = {10, 20, 30, 40, 50};
float marks[3] = {78.5f, 88.0f, 90.2f};
char name[10] = "Hello";
```

### 6.2 Program: Copy one array into another in reverse order
```c
#include <stdio.h>

int main() {
    int arr1[5] = {1, 2, 3, 4, 5};
    int arr2[5];

    for (int i = 0; i < 5; i++) {
        arr2[i] = arr1[4 - i];
    }

    printf("Copied in reverse order:\n");
    for (int i = 0; i < 5; i++) {
        printf("%d ", arr2[i]);
    }
    printf("\n");

    return 0;
}
```

### 6.3 String handling functions
Common string functions:
- `strcpy()` – copy string
- `strcat()` – concatenate string
- `strlen()` – find length
- `strcmp()` – compare strings
- `strrev()` – reverse string

### Example: Concatenate two strings
```c
#include <stdio.h>
#include <string.h>

int main() {
    char s1[20] = "Virat";
    char s2[20] = "Kohli";

    strcat(s1, s2);
    printf("Concatenated string: %s\n", s1);

    return 0;
}
```

### Example: String input/output functions
```c
#include <stdio.h>

int main() {
    char name[30];
    printf("Enter your name: ");
    gets(name);   // old style, avoid in modern C
    printf("Hello, %s\n", name);
    return 0;
}
```

> For modern C, use `fgets()` instead of `gets()`.

---

## 7. Pointers in C

### Important pointer concepts
- Pointer stores address of another variable.
- `*` is dereference operator.
- `&` gives address.

### Example: Basic pointer
```c
#include <stdio.h>

int main() {
    int x = 10;
    int *p = &x;

    printf("Value of x = %d\n", x);
    printf("Address of x = %p\n", (void *)&x);
    printf("Pointer p points to = %d\n", *p);

    return 0;
}
```

### 7.1 Void pointer
A `void *` pointer can point to any data type.

Example:
```c
#include <stdio.h>

int main() {
    int a = 10;
    float b = 5.5f;
    void *vp;

    vp = &a;
    printf("Value of int via void pointer = %d\n", *(int *)vp);

    vp = &b;
    printf("Value of float via void pointer = %.2f\n", *(float *)vp);

    return 0;
}
```

### 7.2 Pointer to pointer
A pointer that stores address of another pointer.

```c
#include <stdio.h>

int main() {
    int x = 50;
    int *p = &x;
    int **pp = &p;

    printf("x = %d\n", x);
    printf("*p = %d\n", *p);
    printf("**pp = %d\n", **pp);

    return 0;
}
```

---

## 8. Structure vs Union

### Structure
- Stores different data types together.
- Each member has its own memory.
- Uses more memory.

### Union
- All members share the same memory location.
- Only one member can be used at a time.
- Saves memory.

### Example: Structure of book details
```c
#include <stdio.h>

struct Book {
    char title[30];
    int price;
    int pages;
};

int main() {
    struct Book b1 = {"C Programming", 250, 400};
    printf("Title: %s\n", b1.title);
    printf("Price: %d\n", b1.price);
    printf("Pages: %d\n", b1.pages);
    return 0;
}
```

### Example: Array of structures
```c
#include <stdio.h>

struct Book {
    char title[30];
    int price;
    int pages;
};

int main() {
    struct Book books[2] = {
        {"C Programming", 250, 400},
        {"Data Structures", 300, 500}
    };

    for (int i = 0; i < 2; i++) {
        printf("Book %d: %s, Price: %d, Pages: %d\n",
               i + 1, books[i].title, books[i].price, books[i].pages);
    }

    return 0;
}
```

### Nested structures
A structure can contain another structure as a member.

Example:
```c
#include <stdio.h>

struct Date {
    int day;
    int month;
    int year;
};

struct Student {
    char name[20];
    struct Date dob;
};

int main() {
    struct Student s = {"Ravi", {10, 12, 2004}};
    printf("Name: %s\n", s.name);
    printf("DOB: %d/%d/%d\n", s.dob.day, s.dob.month, s.dob.year);
    return 0;
}
```

---

## 9. Self-Referential Structure
A self-referential structure contains a pointer to itself.

Use case: linked list, tree nodes.

```c
#include <stdio.h>
#include <stdlib.h>

struct Node {
    int data;
    struct Node *next;
};

int main() {
    struct Node *head = (struct Node *)malloc(sizeof(struct Node));
    head->data = 10;
    head->next = NULL;

    printf("Data: %d\n", head->data);
    free(head);
    return 0;
}
```

---

## 10. File Handling in C

### Important file operations
- `fopen()` – open file
- `fclose()` – close file
- `fscanf()` – read formatted input
- `fprintf()` – write formatted output
- `fgets()` / `fputs()` – string read/write

### Example: Read and display contents of a file
```c
#include <stdio.h>

int main() {
    FILE *fp;
    char ch;

    fp = fopen("sample.txt", "r");
    if (fp == NULL) {
        printf("File cannot be opened\n");
        return 1;
    }

    while ((ch = fgetc(fp)) != EOF) {
        putchar(ch);
    }

    fclose(fp);
    return 0;
}
```

### Example: Copy one file to another and convert lowercase to uppercase
```c
#include <stdio.h>
#include <ctype.h>

int main() {
    FILE *src, *dest;
    char ch;

    src = fopen("input.txt", "r");
    dest = fopen("output.txt", "w");

    if (src == NULL || dest == NULL) {
        printf("File opening failed\n");
        return 1;
    }

    while ((ch = fgetc(src)) != EOF) {
        fputc(toupper(ch), dest);
    }

    fclose(src);
    fclose(dest);
    printf("File copied successfully\n");

    return 0;
}
```

### Command line arguments
`argc` = number of arguments
`argv[]` = actual argument values

Example:
```c
#include <stdio.h>

int main(int argc, char *argv[]) {
    printf("Number of arguments: %d\n", argc);
    for (int i = 0; i < argc; i++) {
        printf("Argument %d: %s\n", i, argv[i]);
    }
    return 0;
}
```

---

## 11. Searching and Time Complexity

### Linear search
Checks each element one by one.
- Best case: $O(1)$
- Worst case: $O(n)$
- Average: $O(n)$

Example:
```c
#include <stdio.h>

int linearSearch(int arr[], int n, int key) {
    for (int i = 0; i < n; i++) {
        if (arr[i] == key)
            return i;
    }
    return -1;
}

int main() {
    int arr[] = {12, 45, 7, 20, 18};
    int key = 20;
    int result = linearSearch(arr, 5, key);

    if (result != -1)
        printf("Found at index %d\n", result);
    else
        printf("Not found\n");

    return 0;
}
```

### Binary search
Works on sorted data only.
- Best case: $O(1)$
- Worst case: $O(log n)$

Example:
```c
#include <stdio.h>

int binarySearch(int arr[], int low, int high, int key) {
    while (low <= high) {
        int mid = low + (high - low) / 2;
        if (arr[mid] == key)
            return mid;
        else if (arr[mid] < key)
            low = mid + 1;
        else
            high = mid - 1;
    }
    return -1;
}

int main() {
    int arr[] = {5, 10, 15, 20, 25};
    int key = 15;
    int result = binarySearch(arr, 0, 4, key);

    if (result != -1)
        printf("Found at index %d\n", result);
    else
        printf("Not found\n");

    return 0;
}
```

---

## 12. Sorting Algorithms

### Quick sort
Quick sort uses a pivot and partitions the array.

### Algorithm idea
1. Pick a pivot
2. Place smaller elements on left and larger on right
3. Recursively sort both partitions

### C implementation
```c
#include <stdio.h>

void swap(int *a, int *b) {
    int temp = *a;
    *a = *b;
    *b = temp;
}

int partition(int arr[], int low, int high) {
    int pivot = arr[high];
    int i = low - 1;

    for (int j = low; j < high; j++) {
        if (arr[j] < pivot) {
            i++;
            swap(&arr[i], &arr[j]);
        }
    }

    swap(&arr[i + 1], &arr[high]);
    return i + 1;
}

void quickSort(int arr[], int low, int high) {
    if (low < high) {
        int pi = partition(arr, low, high);
        quickSort(arr, low, pi - 1);
        quickSort(arr, pi + 1, high);
    }
}

int main() {
    int arr[] = {10, 7, 8, 9, 1, 5};
    int n = sizeof(arr) / sizeof(arr[0]);

    quickSort(arr, 0, n - 1);

    printf("Sorted array: ");
    for (int i = 0; i < n; i++)
        printf("%d ", arr[i]);
    printf("\n");

    return 0;
}
```

### Merge sort
Merge sort divides the array into halves, sorts each half, then merges them.

### Merge sort idea
- Split list into two halves
- Sort both halves recursively
- Merge in order

### Time complexity of merge sort
- Best / average / worst: $O(n \log n)$

---

## 13. Stacks and Queues

### 13.1 Stack concept
A stack is a LIFO (Last In First Out) data structure.

Operations:
- Push
- Pop
- Peek
- isEmpty

### Stack applications
- Function call stack
- Backtracking
- Expression evaluation

### 13.2 Queue concept
A queue is a FIFO (First In First Out) data structure.

Operations:
- Enqueue
- Dequeue
- Front
- Rear

### Queue implementation using arrays
```c
#include <stdio.h>
#define MAX 5

int queue[MAX];
int front = -1, rear = -1;

void enqueue(int value) {
    if (rear == MAX - 1) {
        printf("Queue is full\n");
        return;
    }
    if (front == -1)
        front = 0;
    rear++;
    queue[rear] = value;
}

void dequeue() {
    if (front == -1 || front > rear) {
        printf("Queue is empty\n");
        return;
    }
    printf("Deleted: %d\n", queue[front]);
    front++;
}

void display() {
    for (int i = front; i <= rear; i++)
        printf("%d ", queue[i]);
    printf("\n");
}

int main() {
    enqueue(10);
    enqueue(20);
    enqueue(30);
    display();
    dequeue();
    display();
    return 0;
}
```

---

## 14. Infix to Postfix Conversion

### Important rule
In infix expressions, operators are between operands. In postfix, the operator comes after operands.

Example:
- Infix: `A + B * C`
- Postfix: `A B C * +`

### Stack-based algorithm
1. Read expressions left to right.
2. If operand, print it.
3. If `(`, push to stack.
4. If operator, pop and print higher/equal precedence operators, then push current operator.
5. If `)`, pop until `(`.
6. After scanning, pop remaining operators.

### Example postfix conversion
Input: `A + B * C`
Output: `A B C * +`

---

## 15. Stack using Linked List
A stack can be implemented using a linked list.

### Basic linked-list stack operations
- Push: insert at head
- Pop: remove from head

### Example code
```c
#include <stdio.h>
#include <stdlib.h>

struct Node {
    int data;
    struct Node *next;
};

struct Node *top = NULL;

void push(int value) {
    struct Node *newNode = (struct Node *)malloc(sizeof(struct Node));
    newNode->data = value;
    newNode->next = top;
    top = newNode;
}

void pop() {
    if (top == NULL) {
        printf("Stack is empty\n");
        return;
    }
    struct Node *temp = top;
    printf("Popped: %d\n", temp->data);
    top = top->next;
    free(temp);
}

int main() {
    push(10);
    push(20);
    push(30);
    pop();
    pop();
    return 0;
}
```

---

## 16. Frequently Asked Short Notes

### 16.1 What is an algorithm?
A finite sequence of precise steps to solve a problem.

### 16.2 What is a flowchart?
A graphical representation of an algorithm.

### 16.3 What is the difference between an array and a structure?
- Array: homogeneous data type
- Structure: heterogeneous data type

### 16.4 What is a pointer?
A pointer holds the address of a variable.

### 16.5 What is recursion?
A function calling itself.

### 16.6 What is a union?
A union shares the same memory location among its members.

---

## 17. Expected Most Important Questions for a One-Week Revision

1. Differentiate RAM and ROM.
2. Explain data types in C with examples.
3. Explain control statements in C with examples.
4. Write a factorial program using loops.
5. Explain function definition and function call, with Fibonacci example.
6. Define array and write declaration/initialization examples.
7. Write an array reverse-copy program.
8. Explain string functions and write a concatenation program.
9. Explain pointer, pointer-to-pointer, and void pointer with examples.
10. Differentiate structure and union.
11. Discuss nested structure with example.
12. Explain self-referential structure.
13. Explain file input/output operations and write file-copy code.
14. Explain command line arguments.
15. Define time complexity and compare linear and binary search.
16. Explain quick sort algorithm and give code.
17. Explain stack and queue.
18. Explain infix to postfix conversion.
19. Write stack using linked list.

---

## 18. Exam-Writing Tips

- Always write definitions first, then support with example code.
- For algorithm questions, keep the steps crisp and numbered.
- For coding questions, clearly mention input, processing, and output.
- Always mention time complexity for sorting/searching questions.
- Use simple and clean syntax; avoid unnecessary complexity.

This summary is enough for serious one-week revision of the exam syllabus shown in the past papers.
