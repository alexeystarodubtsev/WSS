import { Component, OnInit } from '@angular/core';
import { DataService } from './data.service';
import { Product } from './product';
import { Phone } from './phone';

@Component({
  selector: 'app',
  templateUrl: './app.component.html',
  providers: [DataService]
})
export class AppComponent implements OnInit {

  product: Product = new Product();   // изменяемый товар
  products: Product[];                // массив товаров
  tableMode: boolean = true;          // табличный режим
  testDoc: File;
  phones: Phone[];
  hasFile: boolean = false;
  constructor(private dataService: DataService) { }

  ngOnInit() {
    this.loadProducts();    // загрузка данных при старте компонента  
  }
  // получаем данные через сервис
  loadProducts() {
    this.dataService.getProducts()
      .subscribe((data: Product[]) => this.products = data);
  }
  // сохранение данных
  save() {
    if (this.product.id == null) {
      this.dataService.createProduct(this.product)
        .subscribe((data: Product) => this.products.push(data));
    } else {
      this.dataService.updateProduct(this.product)
        .subscribe(data => this.loadProducts());
    }
    this.cancel();
  }
  editProduct(p: Product) {
    this.product = p;
  }
  cancel() {
    this.product = new Product();
    this.tableMode = true;
  }
  delete(p: Product) {
    this.dataService.deleteProduct(p.id)
      .subscribe(data => this.loadProducts());
  }
  add() {
    this.cancel();
    this.tableMode = false;
  }
  getPhones() {
    const formData = new FormData();


    formData.append(this.testDoc.name, this.testDoc);

    this.dataService.postFile(formData)
      .subscribe((data: Phone[]) => this.phones = data)
  }
  uploadfile(event) {
    if (event.target.files.length > 0) {
      this.hasFile = true;
    }
    else {
      this.hasFile = false;
    }
    this.testDoc = <File>event.target.files[0];


  }
}
  

