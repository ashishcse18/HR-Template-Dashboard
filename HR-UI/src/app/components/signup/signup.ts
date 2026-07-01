import { Component, inject, signal, Output, EventEmitter } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { AuthService } from '../../services/auth.service';

@Component({
  selector: 'app-signup',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './signup.html',
  styleUrl: './signup.css',
})
export class SignupComponent {
  private readonly authService = inject(AuthService);

  @Output() navigateToLogin = new EventEmitter<void>();

  public name = '';
  public email = '';
  public password = '';
  public loading = signal<boolean>(false);

  public alertMessage = signal<string | null>(null);
  public alertType = signal<'error' | 'success'>('error');

  public onSubmit() {
    if (!this.name || !this.email || !this.password) {
      this.showAlert('All fields are required.', 'error');
      return;
    }

    this.loading.set(true);
    this.alertMessage.set(null);

    this.authService.registerDirect(this.email, this.name, this.password).subscribe({
      next: (res) => {
        this.loading.set(false);
        if (res.success) {
          this.showAlert(res.message, 'success');
          setTimeout(() => {
            this.navigateToLogin.emit();
          }, 1500);
        } else {
          this.showAlert(res.message, 'error');
        }
      },
      error: () => {
        this.loading.set(false);
        this.showAlert('An unexpected error occurred.', 'error');
      },
    });
  }

  private showAlert(msg: string, type: 'error' | 'success') {
    this.alertMessage.set(msg);
    this.alertType.set(type);
  }
}