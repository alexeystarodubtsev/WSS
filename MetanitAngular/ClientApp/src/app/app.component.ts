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
  Attachments: File[];
  phones: Phone[];
  OnlyOnePhones: Phone[];
  hasFile: boolean = false;
  firstMode: boolean = true;
  nameout: string ="";
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
    for (let i = 0; i < this.Attachments.length; i++) {
      formData.append(this.Attachments[i].name, this.Attachments[i])
    }
    formData.set(this.nameout, this.nameout);
    this.dataService.postFile(formData)
      .subscribe((data: JSON) => {
        this.phones = data["item1"];
        this.OnlyOnePhones = data["item2"];
      })
  }
  uploadfile(event) {
    if (event.target.files.length > 0) {
      this.hasFile = true;
    }
    else {
      this.hasFile = false;
    }
    this.Attachments = event.target.files;


  }
  ChangeMode(mode: boolean) {
    this.firstMode = mode;
  }
}
  

